/////////////////////////////////////////////////////////////////////////////////////////////////////////
///
///         AngularSP version 1.0.0.0
///         Created by Ryan Schouten, @shrpntknight, https://angularsp.codeplex.com
///
////////////////////////////////////////////////////////////////////////////////////////////////////////
var angularSP = angular.module('AngularSP', []);
angularSP.service('AngularSPREST', ['$http', '$q', function ($http, $q) {
    //Private Variables
    var self = this;
    var cachedDigests = [];

    //Methods
    this.GetUpdatedDigest = function GetUpdateDigest(webUrl, hostUrl) {
        var deff = $q.defer();
        webUrl = self.SanitizeWebUrl(webUrl);

        var needToAdd = false;
        var digest = self.Support.GetDigestFromCache(webUrl, hostUrl);
        needToAdd = digest === null;
        if (digest != null && digest.digestData != null && digest.digestExpires.getTime() > new Date().getTime()) {
            deff.resolve(digest.digestData);
        }
        else {
            var __REQUESTDIGEST;
            var contextInfoPromise = $http({
                url: webUrl + "_api/contextinfo",
                method: "POST",
                headers: {
                    "Accept": "application/json; odata=verbose"
                }
            }).then(function (data) {
                if (needToAdd) {
                    digest = { digestData: null, digestExpires: null, webUrl: webUrl + hostUrl };
                    cachedDigests.push(digest);
                }
                digest.digestData = self.Support.GetJustTheData(data).GetContextWebInformation;
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
    this.GetItemTypeForListName = function GetItemTypeForListName(name) {
        name = name.replace(/_/g, '_x005f_').replace(/-/g, '');
        return "SP.Data." + name.charAt(0).toUpperCase() + name.split(" ").join("").slice(1) + "ListItem";
    }
    this.GetUrlPrefix = function GetUrlPrefix() {
        if (self.IsSharePointHostedApp) {

        }
    }
    this.Support = {

        Guid: new RegExp("^(\{{0,1}([0-9a-fA-F]){8}-([0-9a-fA-F]){4}-([0-9a-fA-F]){4}-([0-9a-fA-F]){4}-([0-9a-fA-F]){12}\}{0,1})$"),
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
        EndsWith: function endsWith(str, suffix) {
            return str.indexOf(suffix, str.length - suffix.length) !== -1;
        },
        GetCurrentDigestValue: function GetCurrentDigestValue(webUrl, hostUrl) {
            webUrl = self.SanitizeWebUrl(webUrl);
            var digest = self.Support.GetDigestFromCache(webUrl, hostUrl);
            if (digest != null && digest.digestData != null && digest.digestExpires.getTime() > new Date().getTime()) {
                return digest.digestData.FormDigestValue;
            }
            else {
                return $("#__REQUESTDIGEST").val();
            }
        },
        GetDigestFromCache: function GetDigestFromCache(webUrl, hostUrl) {
            for (var i = 0; i < cachedDigests.length; i++) {
                if (cachedDigests[i].webUrl == webUrl + hostUrl)
                    return cachedDigests[i];
            }
            return null;
        },
        SendRequestViaExecutor: function SendRequestViaExecutor(url, appWebUrl, hostUrl, data, method, headers) {
            if (typeof (method) === "undefined" || method === null)
                method = "GET";

            var executor = new SP.RequestExecutor(appWebUrl);

            var context = {
                promise: $q.defer()
            };
            if (url.indexOf("?") > 0)
                url += "&";
            else
                url += "?";

            var requestObj = {
                url:
                    appWebUrl +
                    "_api/SP.AppContextSite(@target)" + url + "@target='" +
                    hostUrl + "'",
                method: method,
                headers: { "Accept": "application/json; odata=verbose" },
                success: Function.createDelegate(context, function (data) {
                    if (data.body === "")
                        data.body = "{}";
                    this.promise.resolve(JSON.parse(data.body));
                }),
                error: Function.createDelegate(context, function (data) {
                    if (data.body === "")
                        data.body = "{}";
                    this.promise.reject(JSON.parse(data.body));
                })
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
    this.SanitizeWebUrl = function SanitizeWebUrl(url) {
        if (typeof (url) == "undefined" || url == null || url == "")
            url = _spPageContextInfo.siteAbsoluteUrl;
        if (url.endsWith("/") === false)
            url += "/";
        return url;
    }
    this.CreateListItem = function CreateListItem(listName, webUrl, item, hostUrl) {
        var itemType = self.GetItemTypeForListName(listName);
        var url = "/web/lists/getbytitle('" + listName + "')/items";
        item["__metadata"] = { "type": itemType };
        webUrl = self.SanitizeWebUrl(webUrl);

        var deff = $q.defer();
        $q.when(self.GetUpdatedDigest(webUrl)).then(function () {
            var promise;
            if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
                promise = $http({
                    url: webUrl + "_api" + url,
                    method: "POST",
                    data: item,
                    headers: {
                        "Accept": "application/json;odata=verbose",
                        'Content-Type': 'application/json;odata=verbose',
                        "X-RequestDigest": self.Support.GetCurrentDigestValue(webUrl)
                    }
                });
            }
            else {
                promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, item, "POST");
            }
            promise.then(function (data) { deff.resolve(self.Support.GetJustTheData(data)) }, function (data) { deff.reject(data) });
        });
        return deff.promise;
    }

    this.GetItemById = function GetItemById(itemId, listName, webUrl, extraParams, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);

        var url = "/web/lists";
        if (self.Support.Guid.test(listName)) {
            listName = listName.replace(/\{|\}/gi, "");
            url += "(guid'" + listName + "')"
        }
        else {
            url += "/getbytitle('" + listName + "')";
        }
        url += "/items(" + itemId + ")";

        if (typeof (extraParams) != "undefined" && extraParams != "") {
            url += "?" + extraParams;
        }

        var deff = $q.defer();
        var promise;
        if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
            url = webUrl + "_api" + url;
            promise = $http({
                url: url,
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose" }
            });
        }
        else {
            promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, null, "GET");
        }

        promise.then(function (data) { deff.resolve(self.Support.GetJustTheData(data)) }, function (data) { deff.reject(data) });
        return deff.promise;

        return promise;
    }
    this.GetListItems = function GetListItems(listName, webUrl, options, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);

        if (typeof (options) === "string")
            options = { $filter: options };

        var url = "/web/lists";
        if (self.Support.Guid.test(listName)) {
            listName = listName.replace(/\{|\}/gi, "");
            url += "(guid'" + listName + "')"
        }
        else {
            url += "/getbytitle('" + listName + "')";
        }
        url += "/items";

        if (typeof (options) !== "undefined") {
            var odata = "";
            for (var property in options) {
                if (options.hasOwnProperty(property)) {
                    if (property === "LoadPage") {
                        url = options[property];
                        break;
                    }
                    if (odata.length == 0)
                        odata = "?";
                    odata += property + "=" + options[property] + "&";
                }
            }
            if (odata.lastIndexOf("&") == odata.length - 1) {
                odata = odata.substring(0, odata.length - 1);
            }
            url += odata;
        }
        var promise;
        if (typeof (hostUrl) === "undefined" || hostUrl === null) {
            url = webUrl + "_api" + url;
            promise = $http({
                url: url,
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose" }
            });
        }
        else {
            promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, null, "GET");
        }
        var deff = $q.defer();
        promise.then(function (data) { deff.resolve(self.Support.GetJustTheData(data)) }, function (data) { deff.reject(data) });
        return deff.promise;
    }
    this.GetList = function GetListItems(listName, webUrl, options, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);

        if (typeof (options) === "string")
            options = { $filter: options };

        var url = "/web/lists";
        if (self.Support.Guid.test(listName)) {
            listName = listName.replace(/\{|\}/gi, "");
            url += "(guid'" + listName + "')"
        }
        else {
            url += "/getbytitle('" + listName + "')";
        }

        if (typeof (options) !== "undefined") {
            var odata = "";
            for (var property in options) {
                if (options.hasOwnProperty(property)) {
                    if (property === "LoadPage") {
                        url = options[property];
                        break;
                    }
                    if (odata.length == 0)
                        odata = "?";
                    odata += property + "=" + options[property] + "&";
                }
            }
            if (odata.lastIndexOf("&") == odata.length - 1) {
                odata = odata.substring(0, odata.length - 1);
            }
            url += odata;
        }
        var promise;
        if (typeof (hostUrl) === "undefined" || hostUrl === null) {
            url = webUrl + "_api" + url;
            promise = $http({
                url: url,
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose" }
            });
        }
        else {
            promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, null, "GET");
        }
        var deff = $q.defer();
        promise.then(function (data) { deff.resolve(self.Support.GetJustTheData(data)) }, function (data) { deff.reject(data) });
        return deff.promise;
    }
    this.GetListItemsByCAML = function GetListItemsByCAML(listName, webUrl, camlQuery, options, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);

        var url = "/web/lists";
        if (self.Support.Guid.test(listName)) {
            listName = listName.replace(/\{|\}/gi, "");
            url += "(guid'" + listName + "')"
        }
        else {
            url += "/getbytitle('" + listName + "')";
        }
        url += "/GetItems(query=@v1)?@v1={\"ViewXml\":\"" + camlQuery + "\"}";

        if (typeof (options) !== "undefined") {
            var odata = "";
            for (var property in options) {
                if (options.hasOwnProperty(property)) {
                    if (property === "LoadPage") {
                        url = options[property];
                        break;
                    }
                    if (odata.length == 0)
                        odata = "?";
                    odata += property + "=" + options[property] + "&";
                }
            }
            if (odata.lastIndexOf("&") == odata.length - 1) {
                odata = odata.substring(0, odata.length - 1);
            }
            url += odata;
        }
        var promise;
        if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
            url = webUrl + "_api" + url;
            promise = $http({
                url: url,
                method: "POST",
                headers: {
                    "Accept": "application/json;odata=verbose",
                    "X-RequestDigest": self.Support.GetCurrentDigestValue(webUrl)
                }
            });
        }
        else {
            promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, null, "POST");
        }
        var deff = $q.defer();
        promise.then(function (data) { deff.resolve(self.Support.GetJustTheData(data)) }, function (data) { deff.reject(data) })
        return deff.promise;
    }
    this.UpdateListItem = function UpdateListItem(itemId, listName, webUrl, updateData, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var itemType = self.GetItemTypeForListName(listName);

        var deff = $q.defer();
        $q.when(self.GetUpdatedDigest(webUrl)).then(function () {
            self.GetItemById(itemId, listName, webUrl, null, hostUrl).then(function (data) {
                updateData.__metadata = { "type": data.__metadata.type };
                var promise;
                if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
                    promise = $http({
                        url: data.__metadata.uri,
                        method: "POST",
                        data: JSON.stringify(updateData),
                        headers: {
                            "Accept": "application/json;odata=verbose",
                            'Content-Type': 'application/json;odata=verbose',
                            "X-RequestDigest": self.Support.GetCurrentDigestValue(webUrl),
                            "X-HTTP-Method": "MERGE",
                            "If-Match": data.__metadata.etag
                        }
                    });
                }
                else {
                    var url = "/web/lists/getbytitle('" + listName + "')/items(" + itemId + ")";

                    var headers = {
                        "X-HTTP-Method": "MERGE",
                        "If-Match": data.__metadata.etag
                    };
                    promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, updateData, "POST", headers);
                }

                promise.then(function (data1) { deff.resolve(data1) }, function (data1) { deff.reject(data1) });
            });
        });
        return deff.promise;
    }
    this.DeleteListItem = function DeleteListItem(itemId, listName, webUrl, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var deff = $q.defer();
        $q.when(self.GetUpdatedDigest(webUrl)).then(function () {
            self.GetItemById(itemId, listName, webUrl, null, hostUrl).then(function (data) {
                var promise;
                if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
                    promise = $http({
                        url: data.__metadata.uri,
                        method: "DELETE",
                        headers: {
                            "Accept": "application/json;odata=verbose",
                            "X-Http-Method": "DELETE",
                            "X-RequestDigest": self.Support.GetCurrentDigestValue(webUrl),
                            "If-Match": data.__metadata.etag
                        }
                    });
                }
                else {
                    var headers = {
                        "X-HTTP-Method": "DELETE",
                        "If-Match": data.__metadata.etag
                    };
                    var url = "/web/lists/getbytitle('" + listName + "')/items(" + itemId + ")";
                    promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, null, "DELETE", headers);
                }

                promise.then(function (data1) { deff.resolve(self.Support.GetJustTheData(data1)) }, function (data1) { deff.reject(data1) });

            });
        });
        return deff.promise;
    }
    this.GetGroup = function GetGroup(groupName, includeMembers, webUrl, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);

        var url = "/web/sitegroups?$filter=(Title%20eq%20%27" + groupName + "%27)";
        if (includeMembers)
            url = url + "&$expand=Users";
        var deff = $q.defer();
        var promise;
        if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
            url = webUrl + "_api" + url;
            promise = $http({
                url: url,
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose" }
            });
        }
        else {
            promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, null, "GET");
        }
        promise.then(function (data1) { deff.resolve(self.Support.GetJustTheData(data1)) }, function (data1) { deff.reject(data1) });
        return deff.promise;
    }
    this.GetSiteUsers = function GetSiteUsers(webUrl, options, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);

        var url = "/web/SiteUsers";
        if (typeof (options) !== "undefined") {
            var odata = "";
            for (var property in options) {
                if (options.hasOwnProperty(property)) {
                    if (property === "LoadPage") {
                        url = options[property];
                        break;
                    }
                    if (odata.length == 0)
                        odata = "?";
                    odata += property + "=" + options[property] + "&";
                }
            }
            if (odata.lastIndexOf("&") == odata.length - 1) {
                odata = odata.substring(0, odata.length - 1);
            }
            url += odata;
        }
        var deff = $q.defer();
        var promise;
        if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
            url = webUrl + "_api" + url;
            promise = $http({
                url: url,
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose" }
            });
        }
        else {
            promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, null, "GET");
        }

        promise.then(function (data1) { deff.resolve(self.Support.GetJustTheData(data1)) }, function (data1) { deff.reject(data1) });
        return deff.promise;
    }
    this.GetListViews = function GetListViews(listName, webUrl, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);

        var url = "/web/lists";
        if (self.Support.Guid.test(listName)) {
            listName = listName.replace(/\{|\}/gi, "");
            url += "(guid'" + listName + "')"
        }
        else {
            url += "/getbytitle('" + listName + "')";
        }
        url += "/Views?$expand=ViewFields";

        var deff = $q.defer();
        var promise;
        if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
            url = webUrl + "_api" + url;

            promise = $http({
                url: url,
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose" }
            });
        }
        else {
            promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, null, "GET");
        }

        promise.then(function (data) { deff.resolve(self.Support.GetJustTheData(data)) }, function (data) { deff.reject(data) });
        return deff.promise;

        return promise;
    }
    this.GetDefaultViewByList = function GetDefaultViewByList(listName, webUrl, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);

        var url = "/web/lists";
        if (self.Support.Guid.test(listName)) {
            listName = listName.replace(/\{|\}/gi, "");
            url += "(guid'" + listName + "')"
        }
        else {
            url += "/getbytitle('" + listName + "')";
        }
        url += "/Views?$expand=ViewFields&$filter=DefaultView eq true";

        var deff = $q.defer();
        var promise;
        if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
            url = webUrl + "_api" + url;

            promise = $http({
                url: url,
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose" }
            });
        }
        else {
            promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, null, "GET");
        }

        promise.then(function (data) {
            deff.resolve(self.Support.GetJustTheData(data)[0])
        }, function (data) {
            deff.reject(data)
        });
        return deff.promise;

        return promise;
    }
    this.GetListFieldProperties = function GetListFieldsProperites(listName, webUrl, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);

        var url = "/web/lists";
        if (self.Support.Guid.test(listName)) {
            listName = listName.replace(/\{|\}/gi, "");
            url += "(guid'" + listName + "')"
        }
        else {
            url += "/getbytitle('" + listName + "')";
        }
        url += "/Fields";

        var deff = $q.defer();
        var promise;
        if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
            url = webUrl + "_api" + url;

            promise = $http({
                url: url,
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose" }
            });
        }
        else {
            promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, null, "GET");
        }

        promise.then(function (data) { deff.resolve(self.Support.GetJustTheData(data)) }, function (data) { deff.reject(data) });
        return deff.promise;

        return promise;
    }
    this.GetGroupsByUserId = function GetGroupsByUserId(userId, webUrl, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var url = webUrl + "_api/Web/GetUserById(" + userId + ")?$expand=Groups";

        var deff = $q.defer();
        var promise;
        if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
            url = webUrl + "_api" + url;

            promise = $http({
                url: url,
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose" }
            });
        }
        else {
            promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, null, "GET");
        }
        promise.then(function (data1) {
            deff.resolve(self.Support.GetJustTheData(data1));
        }, function (data1) {
            deff.reject(data1);
        });
        return deff.promise;
    }
    this.GetUserById = function GetUserById(userId, webUrl, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var url = "/Web/GetUserById(" + userId + ")";
        var deff = $q.defer();
        var promise;
        if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
            url = webUrl + "_api" + url;
            promise = $http({
                url: url,
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose" }
            });
        }
        else {
            promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, null, "GET");
        }
        promise.then(function (data1) { deff.resolve(self.Support.GetJustTheData(data1)) }, function (data1) { deff.reject(data1) });
        return deff.promise;

    }
    this.AddUsertoGroup = function AddUsertoGroup(groupId, loginName, webUrl, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var deff = $q.defer();
        $q.when(self.GetUpdatedDigest(webUrl)).then(function () {
            var item = { LoginName: loginName };
            item["__metadata"] = { "type": "SP.User" };
            webUrl = self.SanitizeWebUrl(webUrl);

            var url = "/web/sitegroups(" + groupId + ")/users";

            var promise;
            if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
                url = webUrl + "_api" + url;
                promise = $http({
                    url: url,
                    method: "POST",
                    data: JSON.stringify(item),
                    headers: {
                        "Accept": "application/json;odata=verbose",
                        'Content-Type': 'application/json;odata=verbose',
                        "X-RequestDigest": self.Support.GetCurrentDigestValue(webUrl)
                    }
                });
            }
            else {
                promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, null, "POST");
            }

            promise.then(function (data1) { deff.resolve(self.Support.GetJustTheData(data1)) }, function (data1) { deff.reject(data1) });
        });
        return deff.promise;
    }
    this.GetUserId = function getUserId(loginName) {
        var deffered = $q.defer();
        var context = new SP.ClientContext.get_current();
        var user = context.get_web().ensureUser(loginName);
        context.load(user);
        context.executeQueryAsync(
             Function.createDelegate(user, function () { deffered.resolve(user); }),
             Function.createDelegate(user, function () { deffered.reject(user, args); })
        );
        return deffered.promise;
    }
    this.CreateSubSite = function CreateSubSite(options, webUrl, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var createData = {
            parameters: {
                '__metadata': {
                    'type': 'SP.WebInfoCreationInformation'
                },
                Url: options.siteUrl,
                Title: options.siteName,
                Description: options.siteDescription,
                Language: 1033,
                WebTemplate: options.siteTemplate,
                UseUniquePermissions: options.uniquePermissions
                //CustomMasterUrl: options.MasterUrl,
                //MasterUrl: options.MasterUrl,
                //EnableMinimalDownload: options.MinimalDownload
            }
        };
        var deffered = $q.defer();
        $q.when(self.GetUpdatedDigest(webUrl)).then(function () {

            // Once we have the form digest value, we can create the subsite
            var url = "/web/webinfos/add";

            var promise;
            if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
                url = webUrl + "_api" + url;
                promise = $http({
                    url: url,
                    type: "POST",
                    headers: {
                        "accept": "application/json;odata=verbose",
                        "content-type": "application/json;odata=verbose",
                        "X-RequestDigest": self.Support.GetCurrentDigestValue(webUrl)
                    },
                    data: JSON.stringify(createData)
                });
            }
            else {
                promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, createData, "POST");
            }

            promise.then(function (data) {
                deffered.resolve(self.Support.GetJustTheData(data));
            });
        });
        return deffered.promise;
    }
    this.GetWebData = function GetWebData(webUrl, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var url = webUrl + "_api/web";

        var deff = $q.defer();
        var promise;
        if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
            promise = $http({
                url: url,
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose" }
            });
        }
        else {
            promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, null, "GET");
        }
        promise.then(function (data1) { deff.resolve(self.Support.GetJustTheData(data1)) }, function (data1) { deff.reject(data1) });
        return deff.promise;
    }
    this.UpdateWebData = function UpdateWebData(webUrl, updateData, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);

        var deff = $q.defer();
        $q.when(self.GetUpdatedDigest(webUrl)).then(function () {
            updateData.__metadata = { "type": "SP.Web" };
            self.GetWebData(webUrl, hostUrl).then(function (data) {
                var promise;
                if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
                    promise = $http({
                        url: data.__metadata.uri,
                        type: "POST",
                        data: JSON.stringify(updateData),
                        headers: {
                            "Accept": "application/json;odata=verbose",
                            'Content-Type': 'application/json;odata=verbose',
                            "X-RequestDigest": self.Support.GetCurrentDigestValue(webUrl),
                            "X-HTTP-Method": "MERGE",
                            "If-Match": data.__metadata.etag
                        }
                    });
                }
                else {
                    var headers = {
                        "X-HTTP-Method": "DELETE",
                        "If-Match": data.__metadata.etag
                    };
                    promise = self.Support.SendRequestViaExecutor(data.__metadata.uri, webUrl, hostUrl, updateData, "POST", headers);
                }
                promise.then(function (data1) { deff.resolve(self.Support.GetJustTheData(data1)) }, function (data1) { deff.reject(data1) });
            });
        });
        return deff.promise;
    }

    this.AddFileToLibrary = function AddFileToLibrary(listName, webUrl, fileName, file, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);

        var url = webUrl + "_api/web/lists/getByTitle(@TargetLibrary)/RootFolder/Files/add(url=@TargetFileName,overwrite='true')?@TargetLibrary='" + listName + "'&@TargetFileName='" + fileName + "'";

        var deff = $q.defer();
        $q.when(self.GetUpdatedDigest(webUrl)).then(function () {
            var promise;

            promise = $http({
                url: url,
                method: "POST",
                transformRequest: [],
                data: file,
                headers: {
                    "Accept": "application/json; odata=verbose",
                    "X-RequestDigest": self.Support.GetCurrentDigestValue(webUrl)
                }
            });
            promise.then(function (data1) { deff.resolve(self.Support.GetJustTheData(data1)) }, function (data1) { deff.reject(data1) });
        });
        return deff.promise;
    }
    this.AddFileToFolderInLibrary = function AddFileToLibrary(webUrl, fileName, folder, file, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var url = webUrl + "_api/web/getfolderbyserverrelativeurl('" + folder + "')/files/add(overwrite=true,url='" +
            fileName + "')";

        var deff = $q.defer();
        $q.when(self.GetUpdatedDigest(webUrl)).then(function () {
            var promise;
            if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
                url = webUrl + "_api" + url;

                promise = $http({
                    url: url,
                    method: "POST",
                    transformRequest: [],
                    data: file,
                    headers: {
                        "Accept": "application/json; odata=verbose",
                        "X-RequestDigest": self.Support.GetCurrentDigestValue(webUrl)
                    }
                });
            }
            else {
                promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, file, "POST");
            }
            promise.then(function (data1) { deff.resolve(self.Support.GetJustTheData(data1)) }, function (data1) { deff.reject(data1) });
        });
        return deff.promise;
    }

    this.GetFilesFromFolder = function GetListItems(webUrl, folder, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var url = webUrl + "_api/web/GetFolderByServerRelativeUrl('" + folder + "')/Files?$expand=Author";

        var deff = $q.defer();
        $q.when(self.GetUpdatedDigest(webUrl)).then(function () {
            var promise;
            if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
                promise = $http({
                    url: url,
                    method: "GET",
                    transformRequest: [],
                    headers: {
                        "Accept": "application/json; odata=verbose",
                        "X-RequestDigest": self.Support.GetCurrentDigestValue(webUrl)
                    }
                });
            }
            else {
                promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, null, "GET");
            }
            promise.then(function (data1) { deff.resolve(self.Support.GetJustTheData(data1)) }, function (data1) { deff.reject(data1) });
        });
        return deff.promise;
    }
    this.DeleteFileFromFolder = function GetListItems(webUrl, file, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var url = webUrl + "_api/web/GetFileByServerRelativeUrl('" + file + "')";

        var deff = $q.defer();
        $q.when(self.GetUpdatedDigest(webUrl)).then(function () {
            var promise;
            if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
                promise = $http({
                    url: url,
                    method: "DELETE",
                    transformRequest: [],
                    headers: {
                        "Accept": "application/json;odata=verbose",
                        "X-Http-Method": "DELETE",
                        "X-RequestDigest": self.Support.GetCurrentDigestValue(webUrl),
                        "IF-MATCH": "*"
                    }
                });
            }
            else {
                promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, null, "DELETE");
            }
            promise.then(function (data1) { deff.resolve(self.Support.GetJustTheData(data1)) }, function (data1) { deff.reject(data1) });
        });
        return deff.promise;
    }

    this.SendEmail = function SendEmail(from, to, body, subject, webUrl, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var url = webUrl + "_api/SP.Utilities.Utility.SendEmail";
        var obj = {
            'properties': {
                '__metadata': { 'type': 'SP.Utilities.EmailProperties' },
                'From': from,
                'To': { 'results': [to] },
                'Body': body,
                'Subject': subject
            }
        };

        var deff = $q.defer();
        $q.when(self.GetUpdatedDigest(webUrl)).then(function () {
            var promise;
            if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
                promise = $http({
                    url: url,
                    method: "POST",
                    transformRequest: [],
                    headers: {
                        "Accept": "application/json;odata=verbose",
                        "X-Http-Method": "POST",
                        "X-RequestDigest": self.Support.GetCurrentDigestValue(webUrl),
                        "IF-MATCH": "*"
                    }
                });
            }
            else {
                promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, obj, "POST");
            }
            promise.then(function (data1) { deff.resolve(self.Support.GetJustTheData(data1)) }, function (data1) { deff.reject(data1) });
        });
        return deff.promise;
    }
    this.DeleteFileFromFolder = function GetListItems(webUrl, file, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var url = webUrl + "_api/web/GetFileByServerRelativeUrl('" + file + "')";

        var deff = $q.defer();
        $q.when(self.GetUpdatedDigest(webUrl)).then(function () {
            var promise;
            if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
                promise = $http({
                    url: url,
                    method: "DELETE",
                    transformRequest: [],
                    headers: {
                        "Accept": "application/json;odata=verbose",
                        "X-Http-Method": "DELETE",
                        "X-RequestDigest": self.Support.GetCurrentDigestValue(webUrl),
                        "IF-MATCH": "*"
                    }
                });
            }
            else {
                promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, null, "DELETE");
            }
            promise.then(function (data1) { deff.resolve(self.Support.GetJustTheData(data1)) }, function (data1) { deff.reject(data1) });
        });
        return deff.promise;
    }

    this.SendEmail = function SendEmail(from, to, body, subject, webUrl, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var url = webUrl + "_api/SP.Utilities.Utility.SendEmail";
        var obj = {
            'properties': {
                '__metadata': { 'type': 'SP.Utilities.EmailProperties' },
                'From': from,
                'To': { 'results': [to] },
                'Body': body,
                'Subject': subject
            }
        };

        var deff = $q.defer();
        $q.when(self.GetUpdatedDigest(webUrl)).then(function () {
            var promise;
            if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
                promise = $http({
                    url: url,
                    method: "POST",
                    transformRequest: [],
                    headers: {
                        "Accept": "application/json;odata=verbose",
                        "X-Http-Method": "POST",
                        "X-RequestDigest": self.Support.GetCurrentDigestValue(webUrl),
                        "IF-MATCH": "*"
                    }
                });
            }
            else {
                promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, obj, "POST");
            }
            promise.then(function (data1) { deff.resolve(self.Support.GetJustTheData(data1)) }, function (data1) { deff.reject(data1) });
        });
        return deff.promise;
    }

    this.APIGet = function APIGet(webUrl, apiUrl, options, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);

        if (typeof (options) === "string")
            options = { $filter: options };

        var url = apiUrl;
        if (typeof (options) !== "undefined") {
            var odata = "";
            for (var property in options) {
                if (options.hasOwnProperty(property)) {
                    if (property === "LoadPage") {
                        url = options[property];
                        break;
                    }
                    if (odata.length == 0)
                        odata = "?";
                    odata += property + "=" + options[property] + "&";
                }
            }
            if (odata.lastIndexOf("&") == odata.length - 1) {
                odata = odata.substring(0, odata.length - 1);
            }
            url += odata;
        }
        var promise;
        if (typeof (hostUrl) === "undefined" || hostUrl === null) {
            url = webUrl + url;
            promise = $http({
                url: url,
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose" }
            });
        }
        else {
            promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, null, "GET");
        }
        var deff = $q.defer();
        promise.then(function (data) { deff.resolve(self.Support.GetJustTheData(data)) }, function (data) { deff.reject(data) });
        return deff.promise;
    }
    this.APIWithData = function APIWithData(webUrl, apiUrl, options, data, hostUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var itemType = self.GetItemTypeForListName(listName);

        var deff = $q.defer();
        $q.when(self.GetUpdatedDigest(webUrl)).then(function () {
            self.GetItemById(itemId, listName, webUrl, null, hostUrl).then(function (data) {
                updateData.__metadata = { "type": data.__metadata.type };
                var promise;
                if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
                    promise = $http({
                        url: data.__metadata.uri,
                        method: "POST",
                        data: JSON.stringify(updateData),
                        headers: {
                            "Accept": "application/json;odata=verbose",
                            'Content-Type': 'application/json;odata=verbose',
                            "X-RequestDigest": self.Support.GetCurrentDigestValue(webUrl),
                            "X-HTTP-Method": "MERGE",
                            "If-Match": data.__metadata.etag
                        }
                    });
                }
                else {
                    var url = "/web/lists/getbytitle('" + listName + "')/items(" + itemId + ")";

                    var headers = {
                        "X-HTTP-Method": "MERGE",
                        "If-Match": data.__metadata.etag
                    };
                    promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, updateData, "POST", headers);
                }

                promise.then(function (data1) { deff.resolve(data1) }, function (data1) { deff.reject(data1) });
            });
        });
        return deff.promise;
    }

    this.Search = {
        Get: function Get(webUrl, options) {
            webUrl = self.SanitizeWebUrl(webUrl);
            var url = webUrl + "_api/search/query";

            var params = "";
            for (var property in options) {
                if (options.hasOwnProperty(property)) {
                    if (params.length == 0)
                        params = "?";
                    params += property + "=";
                    if (typeof (options[property]) === "number" || typeof (options[property]) === "boolean") {
                        params += "" + options[property];
                    }
                    else {
                        params += "'" + options[property] + "'";
                    }
                    params += "&";
                }
            }
            if (params.lastIndexOf("&") == params.length - 1) {
                params = params.substring(0, params.length - 1);
            }
            url += params;

            var deff = $q.defer();
            var promise = $http({
                url: url,
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose" }
            });
            promise.then(function (data1) { deff.resolve(self.Search.GetReturnFromResponse(data1)) }, function (data1) { deff.reject(data1) });
            return deff.promise;
        },
        Post: function Post(webUrl, options) {

        },
        GetReturnFromResponse: function GetReturnFromResponse(data) {
            var obj = {
                ElapsedTime: data.data.d.query.ElapsedTime,
                PrimaryQueryResult: {
                    RefinementResults: self.Search.GetArrayFromRefiners(data.data.d.query.PrimaryQueryResult.RefinementResults),
                    RelevantResults: self.Search.GetArrayFromResultList(data.data.d.query.PrimaryQueryResult.RelevantResults),
                    SpecialTermResults: self.Search.GetArrayFromResultList(data.data.d.query.PrimaryQueryResult.SpecialTermResults)
                },
                SpellingSuggestion: data.data.d.query.SpellingSuggestion
            };

            return obj;
        },
        GetArrayFromRefiners: function GetArrayFromRefiners(res) {
            if (res === null)
                return null;
            var ret = [];
            for (var i = 0; i < res.Refiners.results.length; i++) {
                var obj = {
                    Name: res.Refiners.results[i].Name,
                    results: res.Refiners.results[i].Entries.results
                };
                ret.push(obj);
            }
            return ret;
        },
        GetArrayFromResultList: function GetArrayFromResultList(res) {
            if (res === null)
                return null;
            var ret = { Results: [], RowCount: res.RowCount };
            for (var i = 0; i < res.Table.Rows.results.length; i++) {
                var obj = res.Table.Rows.results[i];
                var retObj = {};
                for (var j = 0; j < obj.Cells.results.length; j++) {
                    switch (obj.Cells.results[j].ValueType) {
                        case "Edm.Double":
                        case "Edm.Int64":
                            retObj[obj.Cells.results[j].Key] = Number(obj.Cells.results[j].Value);
                            break;
                        case "Edm.DateTime":
                            retObj[obj.Cells.results[j].Key] = new Date(obj.Cells.results[j].Value);
                            break;
                        case "Edm.Boolean":
                            retObj[obj.Cells.results[j].Key] = obj.Cells.results[j].Value === "true";
                            break;
                        default:
                            retObj[obj.Cells.results[j].Key] = obj.Cells.results[j].Value;
                            break;
                    }
                }
                ret.Results.push(retObj);
            }
            return ret;
        }
    };
    this.Profile = {
        GetCurrentUser: function GetCurrentUser(webUrl, options, hostUrl) {
            webUrl = self.SanitizeWebUrl(webUrl);
            var url = "/SP.UserProfiles.PeopleManager/GetMyProperties";
            if (typeof (options) !== "undefined") {
                var odata = "";
                for (var property in options) {
                    if (options.hasOwnProperty(property)) {
                        if (property === "LoadPage") {
                            url = options[property];
                            break;
                        }
                        if (odata.length == 0)
                            odata = "?";
                        odata += property + "=" + options[property] + "&";
                    }
                }
                if (odata.lastIndexOf("&") == odata.length - 1) {
                    odata = odata.substring(0, odata.length - 1);
                }
                url += odata;
            }
            var promise;
            if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
                url = webUrl + "_api" + url;
                promise = $http({
                    url: url,
                    method: "GET",
                    headers: { "Accept": "application/json; odata=verbose" }
                });
            }
            else {
                promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, null, "GET");
            }
            var deff = $q.defer();
            promise.then(function (data) { deff.resolve(self.Support.GetJustTheData(data)) }, function (data) { deff.reject(data) });
            return deff.promise;
        },
        GetForUser: function GetForUser(webUrl, userName, options, hostUrl) {
            webUrl = self.SanitizeWebUrl(webUrl);
            var url = "/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='" + userName + "'";
            if (typeof (options) !== "undefined") {
                var odata = "";
                for (var property in options) {
                    if (options.hasOwnProperty(property)) {
                        if (property === "LoadPage") {
                            url = options[property];
                            break;
                        }
                        if (odata.length == 0)
                            odata = "?";
                        odata += property + "=" + options[property] + "&";
                    }
                }
                if (odata.lastIndexOf("&") == odata.length - 1) {
                    odata = odata.substring(0, odata.length - 1);
                }
                url += odata;
            }
            var promise;
            if (typeof (hostUrl) === "undefined" || hostUrl === null || hostUrl === "") {
                url = webUrl + "_api" + url;
                promise = $http({
                    url: url,
                    method: "GET",
                    headers: { "Accept": "application/json; odata=verbose" }
                });
            }
            else {
                promise = self.Support.SendRequestViaExecutor(url, webUrl, hostUrl, null, "GET");
            }
            var deff = $q.defer();
            promise.then(function (data) { deff.resolve(self.Support.GetJustTheData(data)) }, function (data) { deff.reject(data) });
            return deff.promise;
        }
    };
}]);
angularSP.service('AngularSPCSOM', ['$q', function ($q) {
    var self = this;

    this.GetItemTypeForListName = function GetItemTypeForListName(name) {
        return "SP.Data." + name.charAt(0).toUpperCase() + name.split(" ").join("").slice(1) + "ListItem";
    }
    this.SanitizeWebUrl = function SanitizeWebUrl(url) {
        if (typeof (url) == "undefined" || url == null || url == "")
            url = _spPageContextInfo.siteAbsoluteUrl;
        return url;
    }
    this.CreateListItem = function CreateListItem(listName, webUrl, item) {
        webUrl = self.SanitizeWebUrl(webUrl);

        var clientContext = new SP.ClientContext(webUrl);
        var list = clientContext.get_web().get_lists().getByTitle(listName);

        var createInfo = new SP.ListItemCreationInformation();
        var listItem = list.addItem(createInfo);
        for (var name in item) {
            listItem.set_item(name, item[name]);
        }
        listItem.update();

        var ctx = {
            Context: clientContext,
            List: list,
            ListItem: listItem
        };

        clientContext.load(ctx.ListItem);
        var deff = $q.defer();
        clientContext.executeQueryAsync(
            Function.createDelegate(ctx,
                function () {
                    deff.resolve(ctx.ListItem.get_fieldValues());
                }),
            Function.createDelegate(ctx,
                function (sender, args) {
                    deff.reject(args);
                }));
        return deff.promise;
    }

    this.GetItemById = function GetItemById(itemId, listName, webUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);

        var clientContext = new SP.ClientContext(webUrl);
        var targetList = clientContext.get_web().get_lists().getByTitle(listName);
        var targetListItem = targetList.getItemById(itemId);
        clientContext.load(targetListItem);
        var deff = $q.defer();
        clientContext.executeQueryAsync(
            function () {
                deff.resolve(targetListItem.get_fieldValues());
            },
            function (sender, args) {
                deff.reject(args);
            });

        return deff.promise;
    }
    this.GetArrayFromJSOMEnumerator = function (enumObj) {
        var Enumerator = enumObj.getEnumerator();
        var ret = [];

        while (Enumerator.moveNext()) {
            var obj = Enumerator.get_current();
            ret.push(obj.get_fieldValues());
        }
        return ret;
    }
    this.GetListItems = function GetListItems(listName, webUrl, camlquery) {
        webUrl = self.SanitizeWebUrl(webUrl);

        var clientContext = new SP.ClientContext(webUrl);
        var oList = clientContext.get_web().get_lists().getByTitle(listName);

        var camlQuery = new SP.CamlQuery();
        if (typeof (camlquery) !== "undefined") {
            camlQuery.set_viewXml(camlquery);
        }
        var ctx = {
            Context: clientContext,
            List: oList
        };
        ctx.collListItem = oList.getItems(camlQuery);

        var deff = $q.defer();
        clientContext.load(ctx.collListItem);
        clientContext.executeQueryAsync(
            Function.createDelegate(ctx,
                function () {
                    var ret = self.GetArrayFromJSOMEnumerator(ctx.collListItem);
                    deff.resolve(ret);
                }),
            Function.createDelegate(this,
                function (sender, args) {
                    deff.reject(args);
                })
        );
        return deff.promise;
    }
    this.UpdateListItem = function UpdateListItem(itemId, listName, webUrl, item) {
        webUrl = self.SanitizeWebUrl(webUrl);

        var clientContext = new SP.ClientContext(webUrl);
        var oList = clientContext.get_web().get_lists().getByTitle(listName);

        var listItem = oList.getItemById(itemId);
        for (var name in item) {
            listItem.set_item(name, item[name]);
        }
        listItem.update();

        var ctx = {
            Context: clientContext,
            List: oList,
            ListItem: listItem
        };

        var deff = $q.defer();
        clientContext.executeQueryAsync(
            Function.createDelegate(ctx,
                function () {
                    deff.resolve(ctx.ListItem.get_fieldValues());
                }),
            Function.createDelegate(ctx,
                function (sender, args) {
                    deff.reject(args);
                })
        );

        return deff.promise;
    }
    this.DeleteListItem = function DeleteListItem(itemId, listName, webUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);

        var clientContext = new SP.ClientContext(webUrl);
        var list = clientContext.get_web().get_lists().getByTitle(listName);
        var oListItem = list.getItemById(itemId);
        oListItem.deleteObject();

        var deff = $q.defer();
        clientContext.executeQueryAsync(
            function () {
                deff.resolve();
            },
            function (sender, args) {
                deff.reject(args);
            }
        );
        return deff.promise;
    }
    this.GetGroup = function GetGroup(groupName, includeMembers, webUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);

        throw {
            name: "GetGroup",
            level: "Not Implemented",
            message: "Not Implemented",
            htmlMessage: "Not Implemented",
            toString: function () { return this.name + ": " + this.message; }
        };

        var url = webUrl + "_api/web/sitegroups?$filter=(Title%20eq%20%27" + groupName + "%27)";
        if (includeMembers)
            url = url + "&$expand=Users";
        var promise = $http({
            url: url,
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose" }
        });

        return promise;
    }
    this.GetSiteUsers = function GetSiteUsers(webUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        throw {
            name: "GetSiteUsers",
            level: "Not Implemented",
            message: "Not Implemented",
            htmlMessage: "Not Implemented",
            toString: function () { return this.name + ": " + this.message; }
        };

        var url = webUrl + "_api/web/SiteUsers";
        var promise = $http({
            url: url,
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose" }
        });

        return promise;
    }
    this.GetUserById = function GetUserById(userId, webUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        throw {
            name: "GetUserById",
            level: "Not Implemented",
            message: "Not Implemented",
            htmlMessage: "Not Implemented",
            toString: function () { return this.name + ": " + this.message; }
        };
        var url = webUrl + "_api/Web/GetUserById(" + userId + ")";
        var promise = $http({
            url: url,
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose" }
        });
        return promise;

    }
    this.AddUsertoGroup = function AddUsertoGroup(groupId, loginName, webUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        throw {
            name: "AddUsertoGroup",
            level: "Not Implemented",
            message: "Not Implemented",
            htmlMessage: "Not Implemented",
            toString: function () { return this.name + ": " + this.message; }
        };
        var item = { LoginName: loginName };
        item["__metadata"] = { "type": "SP.User" };
        webUrl = self.SanitizeWebUrl(webUrl);
        var promise = $http({
            url: webUrl + "_api/web/sitegroups(" + groupId + ")/users",
            type: "POST",
            contentType: "application/json;odata=verbose",
            data: JSON.stringify(item),
            headers: {
                "Accept": "application/json;odata=verbose",
                "X-RequestDigest": $("#__REQUESTDIGEST").val()
            }
        });
        return promise;
    }
    this.GetUserId = function GetUserId(loginName, webUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var deff = $q.defer();
        var context = new SP.ClientContext(webUrl);
        var user = context.get_web().ensureUser(loginName);
        context.load(user);
        context.executeQueryAsync(
             function () { deff.resolve(user); },
             function (sender, args) { deff.reject(user, args); }
        );
        return deff.promise;
    }
    this.CreateSubSite = function CreateSubSite(options, webUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        throw {
            name: "CreateSubSite",
            level: "Not Implemented",
            message: "Not Implemented",
            htmlMessage: "Not Implemented",
            toString: function () { return this.name + ": " + this.message; }
        };
        var createData = {
            parameters: {
                '__metadata': {
                    'type': 'SP.WebInfoCreationInformation'
                },
                Url: options.siteUrl,
                Title: options.siteName,
                Description: options.siteDescription,
                Language: 1033,
                WebTemplate: options.siteTemplate,
                UseUniquePermissions: options.uniquePermissions
                //CustomMasterUrl: options.MasterUrl,
                //MasterUrl: options.MasterUrl,
                //EnableMinimalDownload: options.MinimalDownload
            }
        };
        var deffered = $q.defer().promise;
        // Because we don't have the hidden __REQUESTDIGEST variable, we need to ask the server for the FormDigestValue
        var __REQUESTDIGEST;
        var rootUrl = location.protocol + "//" + location.host;

        var contextInfoPromise = $http({
            url: webUrl + "_api/contextinfo",
            method: "POST",
            headers: {
                "Accept": "application/json; odata=verbose"
            },
            success: function (data) {
                __REQUESTDIGEST = data.d.GetContextWebInformation.FormDigestValue;
            },
            error: function (data, errorCode, errorMessage) {
                alert(errorMessage);
            }
        });

        // Once we have the form digest value, we can create the subsite
        $q.when(contextInfoPromise).done(function () {
            $http({
                url: webUrl + "_api/web/webinfos/add",
                type: "POST",
                headers: {
                    "accept": "application/json;odata=verbose",
                    "content-type": "application/json;odata=verbose",
                    "X-RequestDigest": __REQUESTDIGEST
                },
                data: JSON.stringify(createData)
            }).then(function (data) {
                deffered.resolve(data);
            });
        });
        return deffered;
    }
    this.GetWebData = function GetWebData(webUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        throw {
            name: "GetWebData",
            level: "Not Implemented",
            message: "Not Implemented",
            htmlMessage: "Not Implemented",
            toString: function () { return this.name + ": " + this.message; }
        };
        var url = webUrl + "_api/web";

        var promise = $http({
            url: url,
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose" }
        });
        return promise;
    }
    this.UpdateWebData = function UpdateWebData(webUrl, updateData) {
        webUrl = self.SanitizeWebUrl(webUrl);
        throw {
            name: "UpdateWebData",
            level: "Not Implemented",
            message: "Not Implemented",
            htmlMessage: "Not Implemented",
            toString: function () { return this.name + ": " + this.message; }
        };

        var __REQUESTDIGEST;
        var contextInfoPromise = $http({
            url: webUrl + "_api/contextinfo",
            method: "POST",
            headers: {
                "Accept": "application/json; odata=verbose"
            },
            success: function (data) {
                __REQUESTDIGEST = data.d.GetContextWebInformation.FormDigestValue;
            },
            error: function (data, errorCode, errorMessage) {
                alert(errorMessage);
            }
        });
        var deff = $q.defer().promise;
        updateData.__metadata = { "type": "SP.Web" };
        $q.when(contextInfoPromise).done(function () {
            self.GetWebData(webUrl).then(function (data) {
                $http({
                    url: data.d.__metadata.uri,
                    type: "POST",
                    contentType: "application/json;odata=verbose",
                    data: JSON.stringify(updateData),
                    headers: {
                        "Accept": "application/json;odata=verbose",
                        "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                        "X-HTTP-Method": "MERGE",
                        "If-Match": data.d.__metadata.etag
                    }
                }).then(function (data1) { deff.resolve(data1) }, function (data1) { deff.reject(data1) });
            });
        });
        return deff;
    }

    this.Search = function Get(webUrl, options) {
        webUrl = self.SanitizeWebUrl(webUrl);

        var clientContext = new SP.ClientContext(webUrl);

        var keywordQuery = new Microsoft.SharePoint.Client.Search.Query.KeywordQuery(clientContext);
        if (typeof (options.querytext) !== "undefined") {
            keywordQuery.set_queryText(options.querytext);
        }
        if (typeof (options.rowlimit) !== "undefined") {
            keywordQuery.set_rowLimit(options.rowlimit);
        }
        if (typeof (options.trimduplicates) !== "undefined") {
            keywordQuery.set_trimDuplicates(options.trimduplicates);
        }
        if (typeof (options.sortlist) !== "undefined") {
            keywordQuery.set_sortList(options.sortlist);
        }
        if (typeof (options.querytemplate) !== "undefined") {
            keywordQuery.set_queryTemplate(options.querytemplate);
        }
        if (typeof (options.enableinterleaving) !== "undefined") {
            keywordQuery.set_enableInterleaving(options.enableinterleaving);
        }
        if (typeof (options.sourceid) !== "undefined") {
            keywordQuery.set_sourceId(options.sourceid);
        }
        if (typeof (options.rankingmodelid) !== "undefined") {
            keywordQuery.set_rankingModelId(options.rankingmodelid);
        }
        if (typeof (options.startrow) !== "undefined") {
            keywordQuery.set_startRow(options.startrow);
        }
        if (typeof (options.rowsperpage) !== "undefined") {
            keywordQuery.set_rowsPerPage(options.rowsperpage);
        }
        if (typeof (options.selectproperties) !== "undefined") {
            var selectProperties = keywordQuery.get_selectProperties();
            var properties = [].concat(options.selectproperties);
            for (var i = 0; i < properties.length; i++) {
                selectProperties.add(properties[i]);
            }
        }
        if (typeof (options.trimduplicates) !== "undefined") {
            keywordQuery.set_trimDuplicates(options.trimduplicates);
        }
        if (typeof (options.trimduplicates) !== "undefined") {
            keywordQuery.set_trimDuplicates(options.trimduplicates);
        }
        if (typeof (options.trimduplicates) !== "undefined") {
            keywordQuery.set_trimDuplicates(options.trimduplicates);
        }
        if (typeof (options.trimduplicates) !== "undefined") {
            keywordQuery.set_trimDuplicates(options.trimduplicates);
        }
        if (typeof (options.trimduplicates) !== "undefined") {
            keywordQuery.set_trimDuplicates(options.trimduplicates);
        }
        if (typeof (options.trimduplicates) !== "undefined") {
            keywordQuery.set_trimDuplicates(options.trimduplicates);
        }
        if (typeof (options.trimduplicates) !== "undefined") {
            keywordQuery.set_trimDuplicates(options.trimduplicates);
        }
        if (typeof (options.trimduplicates) !== "undefined") {
            keywordQuery.set_trimDuplicates(options.trimduplicates);
        }
        if (typeof (options.trimduplicates) !== "undefined") {
            keywordQuery.set_trimDuplicates(options.trimduplicates);
        }
        if (typeof (options.trimduplicates) !== "undefined") {
            keywordQuery.set_trimDuplicates(options.trimduplicates);
        }
        if (typeof (options.trimduplicates) !== "undefined") {
            keywordQuery.set_trimDuplicates(options.trimduplicates);
        }
        if (typeof (options.trimduplicates) !== "undefined") {
            keywordQuery.set_trimDuplicates(options.trimduplicates);
        }
        if (typeof (options.trimduplicates) !== "undefined") {
            keywordQuery.set_trimDuplicates(options.trimduplicates);
        }
        if (typeof (options.trimduplicates) !== "undefined") {
            keywordQuery.set_trimDuplicates(options.trimduplicates);
        }
        if (typeof (options.trimduplicates) !== "undefined") {
            keywordQuery.set_trimDuplicates(options.trimduplicates);
        }
        if (typeof (options.trimduplicates) !== "undefined") {
            keywordQuery.set_trimDuplicates(options.trimduplicates);
        }
        if (typeof (options.trimduplicates) !== "undefined") {
            keywordQuery.set_trimDuplicates(options.trimduplicates);
        }
        if (typeof (options.trimduplicates) !== "undefined") {
            keywordQuery.set_trimDuplicates(options.trimduplicates);
        }
        if (typeof (options.trimduplicates) !== "undefined") {
            keywordQuery.set_trimDuplicates(options.trimduplicates);
        }

        /*culture
        refiners
        refinementfilters
        hiddenconstraints
        enablestemming
        trimduplicatesincludeid
        timeout
        enablenicknames
        enablephonetic
        enablefql
        hithighlightedproperties
        bypassresulttypes
        processbestbets
        clienttype
        personalizationdata
        resultsurl
        querytag
        enablequeryrules
        enablesorting*/

        var searchExecutor = new Microsoft.SharePoint.Client.Search.Query.SearchExecutor(clientContext);
        var results = searchExecutor.executeQuery(keywordQuery);

        var ctx = {
            results: results
        };
        var deff = $q.defer();
        clientContext.executeQueryAsync(
            Function.createDelegate(ctx,
                function () {
                    var obj = {
                        ElapsedTime: this.results.m_value.ElapsedTime,
                        PrimaryQueryResult: {
                            RefinementResults: [],
                            RelevantResults: [],
                            SpecialTermResults: []
                        },
                        SpellingSuggestion: this.results.m_value.SpellingSuggestion
                    };
                    var results = this.results;
                    $.each(results.m_value.ResultTables, function (index, table) {
                        if (table.TableType == "RelevantResults") {
                            obj.PrimaryQueryResult.RelevantResults = results.m_value.ResultTables[index].ResultRows;
                        }
                        else if (table.TableType == "RefinementResults") {
                            obj.PrimaryQueryResult.RefinementResults = results.m_value.ResultTables[index].ResultRows;
                        }
                        else if (table.TableType == "SpecialTermResults") {
                            obj.PrimaryQueryResult.SpecialTermResults = results.m_value.ResultTables[index].ResultRows;
                        }
                    });

                    deff.resolve(obj);
                }),
            Function.createDelegate(ctx,
                function (sender, args) {
                    deff.reject(args);
                })
        );

        return deff.promise;
    }
}]);