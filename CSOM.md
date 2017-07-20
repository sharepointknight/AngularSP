# CSOM Methods

Below are all the currently implemented methods for access the REST services of SharePoint. All of these methods return a [promise](www.html5rocks.com/en/tutorials/es6/promises/).


### CreateListItem
This method will create a new item in the list specified
{{
CreateListItem(listName, webUrl, item)
}}
**Parameters**
* listName- Name of the list that the iten will be added to.
* webUrl- The URL of the SharePoint site where the desired list exists.
* item- Object with properties for each field to set. Included below is an example of setting the Title and a custom field.
{{
     var obj = { 
          Title: 'Test Title', 
          MyCustomField: 'Testing'
     }
}}

### DeleteListItem
This method will delete an item from a list by ID.
{{
DeleteListItem(itemId, listName, webUrl)
}}
**Parameters**
* itemId- Id of the item to delete
* listName- Name of the list where the item exists
* webUrl- The URL of the SharePoint site where the desired list exists.


### GetItemById
This method will retrieve an item from a list by ID.
{{
GetItemById(itemId, listName, webUrl)
}}
**Parameters**
* itemId- ID of the item to be retreived.
* listName- Name of the list where the item exists
* webUrl- The URL of the SharePoint site where the desired list exists.

### GetListItems
This method will retrieve items from a list by using a CAML query .
{{
GetListItems(listName, webUrl, camlquery)
}}
**Parameters**
* listName- Name of the list where the items exist
* webUrl- The URL of the SharePoint site where the desired list exists.
* camlquery- The CAML Query to send to SharePoint. If you want all items in a list pass in an empty string.

### GetUserId
This method will retrieve the user object from SharePoint based on the login name provided. This call will get the user if they are already part of the site, if not it will call ensure user so that there is a user object to return. This can cause issues if the user that is logged in doesn't have permissions to add users.
{{
GetUserId(loginName, webUrl)
}}
**Parameters**
* loginName- The login of the user to retrieve
* webUrl- The URL of the SharePoint site where the desired user exists.

### UpdateListItem
This method will update a list item with the values you provide.
{{
UpdateListItem(itemId, listName, webUrl, item)
}}
**Parameters**
* itemId- The ID of the item that you want to update.
* listName- Name of the list where the item exists
* webUrl- The URL of the SharePoint site where the desired list exists.
* item- Object with properties for each field to set. Included below is an example of setting the Title and a custom field.
{{
     var obj = { 
          Title: 'Test Title', 
          MyCustomField: 'Testing'
     }
}}



