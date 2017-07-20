# REST Methods

Below are all the currently implemented methods for access the REST services of SharePoint. All of these methods return a [promise](www.html5rocks.com/en/tutorials/es6/promises/).

### AddFileToLibrary
This method will add the file to the desired SharePoint library. This is helpful for HTML 5 file upload scenarios.
{{
AddFileToLibrary(listName, webUrl, fileName, file)
}}
**Parameters**
* listName - Name of the library where the file will be added.
* webUrl - String value of the URL of the site, if null or an empty string is passed the service will get the current context URL from SharePoint's context variable.
* fileName - Name of the file when it is created in SharePoint
* file - The actual file that will be uploaded. This could be acquired through HTML5 file upload methods or a file that was generated in your script.

### AddUsertoGroup
This method will add the specified user to a group.
{{
AddUsertoGroup(groupId, loginName, webUrl)
}}
**Parameters**
* groupId - Numeric ID of the group where the user should be added.
* loginName - String value of the login name of the user to add to the group.
* webUrl - String value of the URL of the site, if null or an empty string is passed the service will get the current context URL from SharePoint's context variable

### CreateListItem
This method will create a list item in the list specified.
{{
CreateListItem(listName, webUrl, item)
}}
**Parameters**
* listName - String value of the name of the list where the item will be created.
* webUrl - String value of the URL of the site to find the list, if null or an empty string is passed the service will get the current context URL from SharePoint's context variable
* item - Object with properties for each field to set. Included below is an example of setting the Title and a custom field.
{{
     var obj = { 
          Title: 'Test Title', 
          MyCustomField: 'Testing'
     }
}}

### CreateSubSite
This method will create a sub site under the specified site.
{{
CreateSubSite(options, webUrl)
}}
**Parameters**
* options - Object with the following properties
{{
     siteUrl: "UrlForSite",
     siteName: "NameOfTheNewSite",
     siteDescription: "DescriptionOfTheSite",
     siteTemplate: "TemplateOfTheNewSite",
     uniquePermissions: true|false
}}
* webUrl - String value of the URL of the parent site, if null or an empty string is passed the service will get the current context URL from SharePoint's context variable

### DeleteListItem
This method will delete a list item by ID from the list specified.
{{
DeleteListItem(itemId, listName, webUrl)
}}
**Parameters**
* itemID - Numeric ID of the list item to delete.
* listName - String value of the name of the list where the item will be deleted.
* webUrl - String value of the URL of the site to find the list, if null or an empty string is passed the service will get the current context URL from SharePoint's context variable

### GetGroup
This method will get a SharePoint group by name.
{{
GetGroup(groupName, includeMembers, webUrl)
}}
**Parameters**
* groupName - Name of the SharePoint Group to retrieve.
* includeMembers - Boolean of whether to also pull group members or just return group information.
* webUrl - String value of the URL of the site to find the list, if null or an empty string is passed the service will get the current context URL from SharePoint's context variable

### GetItemById
This method will get a list item by ID from the list specified.
{{
GetItemById(itemId, listName, webUrl)
}}
**Parameters**
* itemID - Numeric ID of the list item to retreive
* listName - String value of the name of the list where the item will be created.
* webUrl - String value of the URL of the site to find the list, if null or an empty string is passed the service will get the current context URL from SharePoint's context variable

### GetListItems
This method will retrieve list items from the list specified.
{{
GetListItems(listName, webUrl, options)
}}
**Parameters**
* listName - String value of the name of the list where the item will be created.
* webUrl - String value of the URL of the site to find the list, if null or an empty string is passed the service will get the current context URL from SharePoint's context variable
* options - This is an object with the different OData options to pass (This parameter is new since version 0.0.0.2). To get a page while paging pass in an option of "LoadPage" with a value that contains the URL retrieved from the __Next property that is returned after making a call.
{{
     var options = {
          $filter: "Title eq 'Test'"
     }
}}

### GetListItemsByCAML
This method will retrieve list items from the list specified using a CAML query.
{{
GetListItemsByCAML(listName, webUrl, camlQuery, options)
}}
**Parameters**
* listName - String value of the name of the list where the item will be created.
* webUrl - String value of the URL of the site to find the list, if null or an empty string is passed the service will get the current context URL from SharePoint's context variable
* camlQuery - String value of the query to pass to the list.
* options - This is an object with the different OData options to pass (This parameter is new since version 0.0.0.2)
{{
     var options = {
          $filter: "Title eq 'Test'"
     }
}}

### GetSiteUsers
This method is will get all the users on the site.
{{
GetSiteUsers(webUrl)
}}

### GetUpdateDigest
This method will retrieve either the cached digest information or will go get a new digest when needed.
{{
GetUpdateDigest(webUrl)
}}
**Parameters**
* webUrl - String value of the URL of the site, if null or an empty string is passed the service will get the current context URL from SharePoint's context variable

### GetUserById
This method will retrieve a site user based on id.
{{
GetUserById(userId, webUrl)
}}
**Parameters**
* userId - Numeric ID of the user in the site collection.
* webUrl - String value of the URL of the site, if null or an empty string is passed the service will get the current context URL from SharePoint's context variable

### GetWebData
This method will retrieve the properties of the requested site.
{{
GetWebData(webUrl)
}}
**Parameters**
* webUrl - String value of the URL of the site, if null or an empty string is passed the service will get the current context URL from SharePoint's context variable

### SanitizeWebUrl
This method is used to ensure that we have a valid site URL to work with. If a null value or empty string is passed to this method then this will retrieve the site URL from the SharePoint context variable.
{{
SanitizeWebUrl(url)
}}

### UpdateListItem
This method will update the list item in the list specified.
{{
UpdateListItem(itemId, listName, webUrl, updateData)
}}
**Parameters**
* itemId - Numeric ID of the list item to update.
* listName - String value of the name of the list where the item will be updated.
* webUrl - String value of the URL of the site to find the list, if null or an empty string is passed the service will get the current context URL from SharePoint's context variable
* updateData - Object with properties for each field to set. Included below is an example of setting the Title and a custom field.
{{
     var obj = { 
          Title: 'Test Title', 
          MyCustomField: 'Testing'
     }
}}

### UpdateWebData
This method will update the site with the data provided.
{{
UpdateWebData(webUrl, updateData)
}}
**Parameters**
* webUrl - String value of the URL of the site to find the list, if null or an empty string is passed the service will get the current context URL from SharePoint's context variable
* updateData - Object with properties for each property to set. Included below is an example of setting the Title.
{{
     var obj = { 
          Title: 'Test Title'
     }
}}

!!Search
The following methods are a grouping of methods related to using the SharePoint Search service

### Search.Get
This method will update the site with the data provided.
{{
Search.Get(webUrl, options)
}}
**Parameters**
* webUrl - String value of the URL of the site to find the list, if null or an empty string is passed the service will get the current context URL from SharePoint's context variable
* options- Object with properties for each property to set. Included below is an example of getting all results containing the keyword SharePoint.
{{
     var options = {
                    QueryText: 'SharePoint'
                }
}}

!!Profile
The following methods are related to retrieving User data from the Profile Store

### Profile.GetCurrentUser
This method will get the user profile data for the currently logged in user
{{
Profile.GetCurrentUser(webUrl, options)
}}
**Parameters**
* webUrl - String value of the URL of the site to find the list, if null or an empty string is passed the service will get the current context URL from SharePoint's context variable
* options- Object with properties for each property to set. These are for the odata operations that you want as part of your query.

### Profile.GetForUser
This method will get the user profile data for the requested user
{{
Profile.GetForUser(webUrl, userName, options)
}}
**Parameters**
* webUrl - String value of the URL of the site to find the list, if null or an empty string is passed the service will get the current context URL from SharePoint's context variable
* userName - The username of the user whose profile information you would like returned.
* options- Object with properties for each property to set. These are for the odata operations that you want as part of your query.

