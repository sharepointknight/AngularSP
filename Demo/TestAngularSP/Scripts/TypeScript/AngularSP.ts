/// <reference path="angularsp.d.ts" />
/// <reference path="../Typings/angularjs/angular.d.ts" />
/// <reference path="../Typings/microsoft-ajax/microsoft.ajax.d.ts" />
/// <reference path="../Typings/sharepoint/sharepoint.d.ts" />

module AngularSP {
    export class AngularSPRest implements AngularSP.AngularSPRest
    {
        private self: any;
        private cachedDigests = [];
        static $inject = ['$http', '$q'];
        constructor($http: angular.IHttpProvider, $q: angular.IQService) {
            this.self = this;
            this.self.$http = $http;
            this.self.$q = $q;
        }    
        
        Support = {
            GetJustTheData: function GetJustTheData(value) {
                var tmp = value;
                if (typeof (tmp.data) !== "undefined") {
                    tmp = tmp.data;
                }
                if (typeof (tmp.d) !== "undefined") {
                    tmp = tmp.d;
                }
                if (typeof (tmp.results) !== "undefined")
                    tmp = tmp.results;
                return tmp;
            },
            EndsWith: function endsWith(str:string, suffix:string):boolean {
                return str.indexOf(suffix, str.length - suffix.length) !== -1;
            },
            GetCurrentDigestValue: function GetCurrentDigestValue(webUrl:string, hostUrl?:string) {
                webUrl = this.self.SanitizeWebUrl(webUrl);
                var digest = this.self.Support.GetDigestFromCache(webUrl, hostUrl);
                if (digest != null && digest.digestData != null && digest.digestExpires.getTime() > new Date().getTime()) {
                    return digest.digestData.FormDigestValue;
                }
                else {
                    return $("#__REQUESTDIGEST").val();
                }
            },
            GetDigestFromCache: function GetDigestFromCache(webUrl, hostUrl) {
                for (var i = 0; i < this.cachedDigests.length; i++) {
                    if (this.cachedDigests[i].webUrl == webUrl + hostUrl)
                        return this.cachedDigests[i];
                }
                return null;
            },
            SendRequestViaExecutor: function SendRequestViaExecutor(url:string, appWebUrl:string, hostUrl:string, data:any, method:string, headers:any) {
                if (typeof (method) === "undefined" || method === null)
                    method = "GET";

                var executor = new SP.RequestExecutor(appWebUrl);

                var context = {
                    promise: this.self.$q.defer()
                };
                if (url.indexOf("?") > 0)
                    url += "&";
                else
                    url += "?";

                var requestObj: SP.RequestInfo = {
                    url:
                    appWebUrl +
                    "_api/SP.AppContextSite(@target)" + url + "@target='" +
                    hostUrl + "'",
                    method: method,
                    body: null,
                    headers: { "Accept": "application/json; odata=verbose" },
                    success: function (data:SP.ResponseInfo) {
                        if (data.body === "")
                            data.body = "{}";
                        this.promise.resolve(JSON.parse(String.fromCharCode.apply(null, data.body)));
                    },
                    error: function (data:SP.ResponseInfo) {
                        if (data.body === "")
                            data.body = "{}";
                        this.promise.reject(JSON.parse(String.fromCharCode.apply(null, data.body)));
                    }
                };
                if (typeof (headers) !== "undefined") {
                    for (var key in headers) {
                        if (headers.hasOwnProperty(key)) {
                            requestObj.headers[key] = headers[key];
                        }
                    }
                }
                if (typeof (data) !== "undefined" && data != null) {
                    requestObj.body = JSON.stringify(data);
                    requestObj.headers["content-type"] = "application/json;odata=verbose";

                    /*if(!create)
                    {
                        requestObj.headers["If-Match"] = "*";
                        requestObj.headers["X-HTTP-Method"] = "MERGE";
                    }*/
                }
                executor.executeAsync(requestObj);
                return context.promise.promise;
            }
        }
        SanitizeWebUrl(url:string):string {
            if (typeof (url) == "undefined" || url == null || url == "")
                url = _spPageContextInfo.siteAbsoluteUrl;
            if (url.endsWith("/") === false)
                url += "/";
            return url;
        }    

        GetUpdateDigest(webUrl, hostUrl) {
            var deff = this.self.$q.defer();
            webUrl = this.self.SanitizeWebUrl(webUrl);

            var needToAdd = false;
            var digest = this.self.Support.GetDigestFromCache(webUrl, hostUrl);
            needToAdd = digest === null;
            if (digest != null && digest.digestData != null && digest.digestExpires.getTime() > new Date().getTime()) {
                deff.resolve(digest.digestData);
            }
            else {
                var __REQUESTDIGEST;
                var contextInfoPromise = this.self.$http({
                    url: webUrl + "_api/contextinfo",
                    method: "POST",
                    headers: {
                        "Accept": "application/json; odata=verbose"
                    }
                }).then(function (data) {
                    if (needToAdd) {
                        digest = { digestData: null, digestExpires: null, webUrl: webUrl + hostUrl };
                        this.self.cachedDigests.push(digest);
                    }
                    digest.digestData = this.self.Support.GetJustTheData(data).GetContextWebInformation;
                    var timeout = digest.digestData.FormDigestTimeoutSeconds;
                    var now = new Date();
                    digest.digestExpires = new Date(now.getTime() + timeout * 1000);
                    deff.resolve(digest.digestData);
                }, function (sender, args) {
                    console.log("Error getting new digest");
                    deff.reject(args);
                });
            }
            return deff.promise;
        }
        GetItemTypeForListName(name) {
            name = name.replace(/_/g, '_x005f_').replace(/-/g, '');
            return "SP.Data." + name.charAt(0).toUpperCase() + name.split(" ").join("").slice(1) + "ListItem";
        }
    }
}