/////////////////////////////////////////////////////////////////////////////////////////////////////////
///
///         AngularSP version 0.0.0.3
///         Created by Ryan Schouten, @shrpntknight, https://angularsp.codeplex.com
///
////////////////////////////////////////////////////////////////////////////////////////////////////////
var angularSP = angular.module('AngularSP', []);
angularSP.service('AngularSPREST', ['$http', '$q', function ($http, $q) {
    //Private Variables
    var self = this;
    var cachedDigests = [];

    //Methods
    this.GetUpdatedDigest = function GetUpdateDigest(webUrl)
    {
        var deff = $q.defer();
        webUrl = self.SanitizeWebUrl(webUrl);

        var needToAdd = false;
        var digest = self.Support.GetDigestFromCache(webUrl);
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
                if (needToAdd)
                {
                    digest = { digestData: null, digestExpires: null, webUrl: webUrl };
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
        return "SP.Data." + name.charAt(0).toUpperCase() + name.split(" ").join("").slice(1) + "ListItem";
    }
    this.GetUrlPrefix = function GetUrlPrefix() {
        if (self.IsSharePointHostedApp) {

        }
    }
    this.Support = {
        GetJustTheData: function GetJustTheData(value)
        {
            var tmp = value;
            if (typeof (tmp.data) !== "undefined") {
                tmp = tmp.data;
                if (typeof (tmp.d) !== "undefined") {
                    tmp = tmp.d;
                    if (typeof (tmp.results) !== "undefined")
                        tmp = tmp.results;
                }
            }
            return tmp;
        },
        EndsWith: function endsWith(str, suffix) {
            return str.indexOf(suffix, str.length - suffix.length) !== -1;
        },
        GetCurrentDigestValue: function GetCurrentDigestValue(webUrl)
        {
            webUrl = self.SanitizeWebUrl(webUrl);
            var digest = self.Support.GetDigestFromCache(webUrl);
            if (digest != null && digest.digestData != null && digest.digestExpires.getTime() > new Date().getTime())
            {
                return digest.digestData.FormDigestValue;
            }
            else
            {
                return $("#__REQUESTDIGEST").val();
            }
        },
        GetDigestFromCache: function GetDigestFromCache(webUrl)
        {
            for(var i=0;i<cachedDigests.length;i++)
            {
                if (cachedDigests[i].webUrl == webUrl)
                    return cachedDigests[i];
            }
            return null;
        }
    }
    this.SanitizeWebUrl = function SanitizeWebUrl(url) {
        if (typeof (url) == "undefined" || url == null || url == "")
            url = _spPageContextInfo.siteAbsoluteUrl;
        if (url.endsWith("/") === false)
            url += "/";
        return url;
    }
    this.CreateListItem = function CreateListItem(listName, webUrl, item) {
        var itemType = self.GetItemTypeForListName(listName);
        item["__metadata"] = { "type": itemType };
        webUrl = self.SanitizeWebUrl(webUrl);

        var deff = $q.defer();
        $q.when(self.GetUpdatedDigest(webUrl)).then(function ()
        {
            var promise = $http({
                url: webUrl + "_api/web/lists/getbytitle('" + listName + "')/items",
                method: "POST",
                data: item,
                headers: {
                    "Accept": "application/json;odata=verbose",
                    'Content-Type': 'application/json;odata=verbose',
                    "X-RequestDigest": self.Support.GetCurrentDigestValue(webUrl)
                }
            });
            promise.then(function (data) { deff.resolve(self.Support.GetJustTheData(data)) }, function (data) { deff.reject(data) });
        });
        return deff.promise;
    }

    this.GetItemById = function GetItemById(itemId, listName, webUrl, extraParams) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var url = webUrl + "_api/web/lists/getbytitle('" + listName + "')/items(" + itemId + ")";
        if (typeof (extraParams) != "undefined" && extraParams != "") {
            url += "?" + extraParams;
        }

        var deff = $q.defer();
        var promise = $http({
            url: url,
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose" }
        });

        promise.then(function (data) { deff.resolve(self.Support.GetJustTheData(data)) }, function (data) { deff.reject(data) });
        return deff.promise;

        return promise;
    }
    this.GetListItems = function GetListItems(listName, webUrl, options) {
        webUrl = self.SanitizeWebUrl(webUrl);

        if (typeof (options) === "string")
            options = { $filter: options };

        var url = webUrl + "_api/web/lists/getbytitle('" + listName + "')/items";
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
        var promise = $http({
            url: url,
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose" }
        });
        var deff = $q.defer();
        promise.then(function (data) { deff.resolve(self.Support.GetJustTheData(data)) }, function (data) { deff.reject(data) });
        return deff.promise;
    }
    this.GetListItemsByCAML = function GetListItemsByCAML(listName, webUrl, camlQuery, options) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var url = webUrl + "_api/web/lists/getbytitle('" + listName + "')/GetItems(query=@v1)?@v1={\"ViewXml\":\"" + camlQuery + "\"}";
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
        var promise = $http({
            url: url,
            method: "POST",
            headers: {
                "Accept": "application/json;odata=verbose",
                "X-RequestDigest": self.Support.GetCurrentDigestValue(webUrl)
            }
        });
        var deff = $q.defer();
        promise.then(function (data) { deff.resolve(self.Support.GetJustTheData(data)) }, function (data) { deff.reject(data) })
        return deff.promise;
    }
    this.UpdateListItem = function UpdateListItem(itemId, listName, webUrl, updateData) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var itemType = self.GetItemTypeForListName(listName);

        var deff = $q.defer();
        $q.when(self.GetUpdatedDigest(webUrl)).then(function () {
            self.GetItemById(itemId, listName, webUrl).then(function (data) {
                updateData.__metadata = { "type": data.__metadata.type };
                var promise = $http({
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
                promise.then(function (data1) { deff.resolve(data1) }, function (data1) { deff.reject(data1) });
            });
        });
        return deff.promise;
    }
    this.DeleteListItem = function DeleteListItem(itemId, listName, webUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var deff = $q.defer();
        $q.when(self.GetUpdatedDigest(webUrl)).then(function () {
            self.GetItemById(itemId, listName, webUrl).then(function (data) {
                var promise = $http({
                    url: data.__metadata.uri,
                    method: "DELETE",
                    headers: {
                        "Accept": "application/json;odata=verbose",
                        "X-Http-Method": "DELETE",
                        "X-RequestDigest": self.Support.GetCurrentDigestValue(webUrl),
                        "If-Match": data.__metadata.etag
                    }
                });
                promise.then(function (data1) { deff.resolve(self.Support.GetJustTheData(data1)) }, function (data1) { deff.reject(data1) });
            });
        });
        return deff.promise;
    }
    this.GetGroup = function GetGroup(groupName, includeMembers, webUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);

        var url = webUrl + "_api/web/sitegroups?$filter=(Title%20eq%20%27" + groupName + "%27)";
        if (includeMembers)
            url = url + "&$expand=Users";
        var deff = $q.defer();
        var promise = $http({
            url: url,
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose" }
        });
        promise.then(function (data1) { deff.resolve(self.Support.GetJustTheData(data1)) }, function (data1) { deff.reject(data1) });
        return deff.promise;
    }
    this.GetSiteUsers = function GetSiteUsers(webUrl, options) {
        webUrl = self.SanitizeWebUrl(webUrl);

        var url = webUrl + "_api/web/SiteUsers";
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
        var promise = $http({
            url: url,
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose" }
        });

        promise.then(function (data1) { deff.resolve(self.Support.GetJustTheData(data1)) }, function (data1) { deff.reject(data1) });
        return deff.promise;
    }
    this.GetUserById = function GetUserById(userId, webUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var url = webUrl + "_api/Web/GetUserById(" + userId + ")";
        var deff = $q.defer();
        var promise = $http({
            url: url,
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose" }
        });
        promise.then(function (data1) { deff.resolve(self.Support.GetJustTheData(data1)) }, function (data1) { deff.reject(data1) });
        return deff.promise;

    }
    this.AddUsertoGroup = function AddUsertoGroup(groupId, loginName, webUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var deff = $q.defer();
        $q.when(self.GetUpdatedDigest(webUrl)).then(function () {
            var item = { LoginName: loginName };
            item["__metadata"] = { "type": "SP.User" };
            webUrl = self.SanitizeWebUrl(webUrl);
            var promise = $http({
                url: webUrl + "_api/web/sitegroups(" + groupId + ")/users",
                type: "POST",
                data: JSON.stringify(item),
                headers: {
                    "Accept": "application/json;odata=verbose",
                    'Content-Type': 'application/json;odata=verbose',
                    "X-RequestDigest": self.Support.GetCurrentDigestValue(webUrl)
                }
            });
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
    this.CreateSubSite = function CreateSubSite(options, webUrl) {
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
            $http({
                url: webUrl + "_api/web/webinfos/add",
                type: "POST",
                headers: {
                    "accept": "application/json;odata=verbose",
                    "content-type": "application/json;odata=verbose",
                    "X-RequestDigest": self.Support.GetCurrentDigestValue(webUrl)
                },
                data: JSON.stringify(createData)
            }).then(function (data) {
                deffered.resolve(self.Support.GetJustTheData(data));
            });
        });
        return deffered.promise;
    }
    this.GetWebData = function GetWebData(webUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var url = webUrl + "_api/web";

        var deff = $q.defer();
        var promise = $http({
            url: url,
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose" }
        });
        promise.then(function (data1) { deff.resolve(self.Support.GetJustTheData(data1)) }, function (data1) { deff.reject(data1) });
        return deff.promise;
    }
    this.UpdateWebData = function UpdateWebData(webUrl, updateData) {
        webUrl = self.SanitizeWebUrl(webUrl);

        var deff = $q.defer();
        $q.when(self.GetUpdatedDigest(webUrl)).then(function () {
            updateData.__metadata = { "type": "SP.Web" };
            self.GetWebData(webUrl).then(function (data) {
                $http({
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
                }).then(function (data1) { deff.resolve(self.Support.GetJustTheData(data1)) }, function (data1) { deff.reject(data1) });
            });
        });
        return deff.promise;
    }

    this.AddFileToLibrary = function AddFileToLibrary(listName, webUrl, fileName, file) {
        webUrl = self.SanitizeWebUrl(webUrl);

        var url = webUrl + "_api/web/lists/getByTitle(@TargetLibrary)/RootFolder/Files/add(url=@TargetFileName,overwrite='true')?@TargetLibrary='" + listName + "'&@TargetFileName='" + fileName + "'";

        var deff = $q.defer();
        $q.when(self.GetUpdatedDigest(webUrl)).then(function () {
            var promise = $http({
                url: url,
                method: "POST",
                data: file,
                headers: {
                    "Accept": "application/json; odata=verbose",
                    'Content-Type': 'application/json;odata=verbose',
                    "X-RequestDigest": self.Support.GetCurrentDigestValue(webUrl)
                }
            });
            promise.then(function (data1) { deff.resolve(self.Support.GetJustTheData(data1)) }, function (data1) { deff.reject(data1) });
        });
        return deff.promise;
    }

    this.Search = {
        Get: function Get(webUrl, options)
        {
            webUrl = self.SanitizeWebUrl(webUrl);
            var url = webUrl + "_api/search/query";

            var params = "";
            for (var property in options) {
                if (options.hasOwnProperty(property)) {
                    if (params.length == 0)
                        params = "?";
                    params += property + "=";
                    if (typeof (options[property]) === "number" || typeof (options[property]) === "boolean")
                    {
                        params += "" + options[property];
                    }
                    else
                    {
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
        Post: function Post(webUrl, options)
        {

        },
        GetReturnFromResponse: function GetReturnFromResponse(data)
        {
            var obj = {
                ElapsedTime: data.data.d.query.ElapsedTime,
                PrimaryQueryResult: {
                    RefinementResults: self.Search.GetArrayFromResultList(data.data.d.query.PrimaryQueryResult.RefinementResults),
                    RelevantResults: self.Search.GetArrayFromResultList(data.data.d.query.PrimaryQueryResult.RelevantResults),
                    SpecialTermResults: self.Search.GetArrayFromResultList(data.data.d.query.PrimaryQueryResult.SpecialTermResults)
                },
                SpellingSuggestion: data.data.d.query.SpellingSuggestion
            };

            return obj;
        },
        GetArrayFromResultList: function GetArrayFromResultList(res)
        {
            if (res === null)
                return null;
            var ret = { Results: [], RowCount: res.RowCount };
            for(var i=0;i<res.Table.Rows.results.length;i++)
            {
                var obj = res.Table.Rows.results[i];
                var retObj = {};
                for(var j=0;j<obj.Cells.results.length;j++)
                {
                    switch(obj.Cells.results[j].ValueType)
                    {
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
    this.GetArrayFromJSOMEnumerator = function (enumObj)
    {
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
        if (typeof (camlquery) !== "undefined")
        {
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
    this.GetUserId = function getUserId(loginName, webUrl) {
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
}]);

