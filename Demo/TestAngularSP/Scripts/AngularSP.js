/////////////////////////////////////////////////////////////////////////////////////////////////////////
///
///         AngularSP version 0.0.0.1
///         Created by Ryan Schouten, @shrpntknight, https://github.com/sharepointknight/AngularSP
///
////////////////////////////////////////////////////////////////////////////////////////////////////////
var angularSP = angular.module('AngularSP', []);
angularSP.service('AngularSPREST', ['$http', '$q', function ($http, $q) {
    var self = this;
    this.SetDigestExpiration = function SetDigestExpiration(ticks)
    {

    }
    this.DigestExpires = null;
    this.GetUpdatedDigest = function GetUpdateDigest(webUrl)
    {
        webUrl = self.SanitizeWebUrl(webUrl);
        var promise = $http({
            url: webUrl + "/_api/contextinfo",
            method: "POST",
            headers: {
                "Accept": "application/json;odata=verbose"
            }
        });
    }
    this.GetItemTypeForListName = function GetItemTypeForListName(name) {
        return "SP.Data." + name.charAt(0).toUpperCase() + name.split(" ").join("").slice(1) + "ListItem";
    }
    this.GetUrlPrefix = function GetUrlPrefix() {
        if (self.IsSharePointHostedApp) {

        }
    }
    this.SanitizeWebUrl = function SanitizeWebUrl(url) {
        if (typeof (url) == "undefined" || url == null || url == "")
            url = _spPageContextInfo.siteAbsoluteUrl;
        return url;
    }
    this.CreateListItem = function CreateListItem(listName, webUrl, item) {
        var itemType = self.GetItemTypeForListName(listName);
        //item["__metadata"] = { "type": itemType };
        webUrl = self.SanitizeWebUrl(webUrl);

        var deff = $q.defer();
        var promise = $http({
            url: webUrl + "/_api/web/lists/getbytitle('" + listName + "')/items",
            method: "POST",
            data: item,
            headers: {
                "Accept": "application/json;odata=verbose",
                "X-RequestDigest": $("#__REQUESTDIGEST").val()
            }
        });
        promise.then(function (data) { deff.resolve(data.data.d) }, function (data) { deff.reject(data) });
        return deff.promise;
    }

    this.GetItemById = function GetItemById(itemId, listName, webUrl, extraParams) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var url = webUrl + "/_api/web/lists/getbytitle('" + listName + "')/items(" + itemId + ")";
        if (typeof (extraParams) != "undefined" && extraParams != "") {
            url += "?" + extraParams;
        }
        var promise = $http({
            url: url,
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose" }
        });

        return promise;
    }
    this.GetListItems = function GetListItems(listName, webUrl, filter, sort) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var url = webUrl + "/_api/web/lists/getbytitle('" + listName + "')/items";
        if (typeof (filter) != "undefined" && filter.length > 0) {
            url = url + "?$filter=" + filter;
        }
        if (typeof (sort) != "undefined" && sort.length > 0) {
            if (url.indexOf("?") > 0) {
                url = url + "&";
            }
            else {
                url = url + "?";
            }
            url = url + "$orderby=" + sort;
        }
        var promise = $http({
            url: url,
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose" }
        });
        var deff = $q.defer();
        promise.then(function (data) { deff.resolve(data.data.d.results) }, function (data) { deff.reject(data) });
        return deff.promise;
    }
    this.GetListItemsByCAML = function GetListItemsByCAML(listName, webUrl, camlQuery, extraUrl, extraData) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var url = webUrl + "/_api/web/lists/getbytitle('" + listName + "')/GetItems(query=@v1)?@v1={\"ViewXml\":\"" + camlQuery + "\"}";
        if (extraUrl.length > 0) {
            url += "&" + extraUrl;
        }
        var promise = $http({
            url: url,
            method: "POST",
            headers: {
                "Accept": "application/json;odata=verbose",
                "X-RequestDigest": $("#__REQUESTDIGEST").val()
            }
        });
        var deff = $q.defer();
        promise.then(function (data) { data.ExtraData = extraData; deff.resolve(data) }, function (data) { data.ExtraData = extraData; deff.reject(data) })
        return deff.promise;
    }
    this.UpdateListItem = function UpdateListItem(itemId, listName, webUrl, updateData) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var itemType = self.GetItemTypeForListName(listName);

        //var item = {
        //    "__metadata": { "type": itemType },
        //    "Title": title
        //};

        //updateData.__metadata = { "type": itemType };
        var deff = $q.defer();
        self.GetItemById(itemId, listName, webUrl).then(function (data) {
            //updateData.__metadata = { "type": data.data.d.__metadata.type };
            var promise = $http({
                url: data.data.d.__metadata.uri,
                method: "POST",
                contentType: "application/json;odata=verbose",
                data: JSON.stringify(updateData),
                headers: {
                    "Accept": "application/json;odata=verbose",
                    "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                    "X-HTTP-Method": "MERGE",
                    "If-Match": data.data.d.__metadata.etag
                }
            });
            promise.then(function (data1) { deff.resolve(data1) }, function (data1) { deff.reject(data1) });
        });
        return deff.promise;
    }
    this.DeleteListItem = function DeleteListItem(itemId, listName, webUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var deff = $q.defer();
        self.GetItemById(itemId, listName, webUrl).then(function (data) {
            var promise = $http({
                url: data.data.d.__metadata.uri,
                method: "DELETE",
                headers: {
                    "Accept": "application/json;odata=verbose",
                    "X-Http-Method": "DELETE",
                    "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                    "If-Match": data.data.d.__metadata.etag
                }
            });
            promise.then(function (data1) { deff.resolve(data1) }, function (data1) { deff.reject(data1) });
        });
        return deff.promise;
    }
    this.GetGroup = function GetGroup(groupName, includeMembers, webUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);

        var url = webUrl + "/_api/web/sitegroups?$filter=(Title%20eq%20%27" + groupName + "%27)";
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

        var url = webUrl + "/_api/web/SiteUsers";
        var promise = $http({
            url: url,
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose" }
        });

        return promise;
    }
    this.GetUserById = function GetUserById(userId, webUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var url = webUrl + "/_api/Web/GetUserById(" + userId + ")";
        var promise = $http({
            url: url,
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose" }
        });
        return promise;

    }
    this.AddUsertoGroup = function AddUsertoGroup(groupId, loginName, webUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var item = { LoginName: loginName };
        item["__metadata"] = { "type": "SP.User" };
        webUrl = self.SanitizeWebUrl(webUrl);
        var promise = $http({
            url: webUrl + "/_api/web/sitegroups(" + groupId + ")/users",
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
    this.GetUserId = function getUserId(loginName) {
        var deffered = $q.defer();
        var context = new SP.ClientContext.get_current();
        var user = context.get_web().ensureUser(loginName);
        context.load(user);
        context.executeQueryAsync(
             Function.createDelegate(null, function () { deffered.resolve(user); }),
             Function.createDelegate(null, function () { deffered.reject(user, args); })
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
        // Because we don't have the hidden __REQUESTDIGEST variable, we need to ask the server for the FormDigestValue
        var __REQUESTDIGEST;
        var rootUrl = location.protocol + "//" + location.host;

        var contextInfoPromise = $http({
            url: webUrl + "/_api/contextinfo",
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
                url: webUrl + "/_api/web/webinfos/add",
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
        return deffered.promise;
    }
    this.GetWebData = function GetWebData(webUrl) {
        webUrl = self.SanitizeWebUrl(webUrl);
        var url = webUrl + "/_api/web";

        var promise = $http({
            url: url,
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose" }
        });
        return promise;
    }
    this.UpdateWebData = function UpdateWebData(webUrl, updateData) {
        webUrl = self.SanitizeWebUrl(webUrl);

        var __REQUESTDIGEST;
        var contextInfoPromise = $http({
            url: webUrl + "/_api/contextinfo",
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
        var deff = $q.defer();
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
        return deff.promise;
    }
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
                deff.resolve(targetListItem);
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
    this.GetListItems = function GetListItems(listName, webUrl, camlQuery) {
        webUrl = self.SanitizeWebUrl(webUrl);

        var clientContext = new SP.ClientContext(webUrl);
        var oList = clientContext.get_web().get_lists().getByTitle(listName);

        var camlQuery = new SP.CamlQuery();
        if (typeof (camlQuery) == "undefined")
        {
            camlQuery.set_viewXml(camlQuery);
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

        var url = webUrl + "/_api/web/sitegroups?$filter=(Title%20eq%20%27" + groupName + "%27)";
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

        var url = webUrl + "/_api/web/SiteUsers";
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
        var url = webUrl + "/_api/Web/GetUserById(" + userId + ")";
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
            url: webUrl + "/_api/web/sitegroups(" + groupId + ")/users",
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
            url: webUrl + "/_api/contextinfo",
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
                url: webUrl + "/_api/web/webinfos/add",
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
        var url = webUrl + "/_api/web";

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
            url: webUrl + "/_api/contextinfo",
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

