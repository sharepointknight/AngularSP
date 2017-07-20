# Getting Started

The goal of this project was to make it easy to integrate AngularJS and SharePoint. To this end it is really easy to get started. Just follow the steps below and you will be up and running in no time.

1. Load AngularJS
----
On your page where you will be loading your app reference the AngularJS javascript files.
{{
     <script type="text/javascript" src="../Scripts/angularjs.js"></script>
}}

2. Load AngularSP
----
Also on your page that will be running your app reference the AngularSP javascript file.
{{
     <script type="text/javascript" src="../Content/js/AngularSP.min.js"></script>
}}

3. Create your App and Controllers
----
In order for AngularSP to work you need to add it as a dependency of your App. This will makes the services available to your App.
{{
     angular.module('AngularSPApp', "['AngularSP']('AngularSP'));
}}
Now that AngularSP has been loaded you can reference it in your controller. 
{{
     App.controller('MyController', ['$scope', 'AngularSPCSOM',
     function ($scope, $angularSPCSOM) {
     }]);
}}
Here is the current list of services that we have implemented.
* AngularSPREST
* AngularSPCSOM
We will be implementing at least one more in the near future for the SharePoint Request Executor to handle cross domain REST calls.

4. Call the methods to get the data you want
----
The beauty of this project is how easy it is to access the SharePoint Services. For the most part the calls are very similar between the REST and CSOM variations. Sometimes there will be differences and these will be noted in the documentation for the [Methods](Methods).

Here is an example of getting the items from a list:
{{
     $angularSPCSOM.GetListItems("TestList","SiteUrl").then(function (data) 
     {
          //Work with the data
     });
}}
The data is returned as an array of the items so you don't have to worry about all where in all the sub objects the data resides. The first version is not super consistent yet on this aspect but will be fixed in the next release.