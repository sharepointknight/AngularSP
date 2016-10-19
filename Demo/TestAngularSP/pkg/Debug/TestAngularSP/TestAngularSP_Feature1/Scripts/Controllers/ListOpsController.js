'use strict';

SPKnight.TestApp.App.controller('ListOpsController', ['$scope', '$rootScope', '$filter', 'AngularSPREST', 'AngularSPCSOM',
    function ($scope, $rootScope, $filter, $angularSPRest, $angularSPCSOM) {
        $scope.Items = [];
        $scope.Method = "REST";
        $scope.NewItemTitle = "";
        $scope.ListUrl = "";
        $scope.ListName = "";

        $scope.GetListItemsHost = function GetListItemsHost()
        {
            $angularSPRest.GetListItems($scope.ListName, "/TestAngularSP", null, $scope.ListUrl).then(function (items) {
                $scope.Items = items;
            });
        }
        $scope.GetListItems = function GetListItems()
        {
            if($scope.Method === "REST")
            {
                $angularSPRest.GetListItems("TestList", "/TestAngularSP").then(function (items) {
                    $scope.Items = items;
                });
            }
            else
            {
                var res = $angularSPCSOM.GetListItems("TestList", "/TestAngularSP");
                res.then(function (items) {
                    $scope.Items = items;
                });
            }
        }
        $scope.UpdateItem = function UpdateItem(item)
        {
            if ($scope.Method === "REST") {
                $angularSPRest.UpdateListItem(item.ID, "TestList", "/TestAngularSP", { Title: item.Title }, $scope.ListUrl).then(function (res) {
                    debugger;
                });
            }
            else {
                $angularSPCSOM.UpdateListItem(item.ID, "TestList", "/TestAngularSP", { Title: item.Title }).then(function (res) {
                    debugger;
                });
            }
        }
        $scope.DeleteItem = function DeleteItem(item)
        {
            if ($scope.Method === "REST") {
                $angularSPRest.DeleteListItem(item.ID, "TestList", "/TestAngularSP", $scope.ListUrl).then(function (res) {
                    var index = $scope.Items.indexOf(item);
                    $scope.Items.splice(index, 1);
                });
            }
            else {
                $angularSPCSOM.DeleteListItem(item.ID, "TestList", "/TestAngularSP").then(function (res) {
                    var index = $scope.Items.indexOf(item);
                    $scope.Items.splice(index, 1);
                });
            }
        }
        $scope.CreateItem = function CreateItem()
        {
            var item = {
                Title: $scope.NewItemTitle
            };
            if ($scope.Method === "REST") {
                $angularSPRest.CreateListItem("TestList", "/TestAngularSP", item, $scope.ListUrl).then(function (res) {
                    $scope.NewItemTitle = "";
                    $scope.Items.push(res);
                });
            }
            else {
                $angularSPCSOM.CreateListItem("TestList", "/TestAngularSP", item).then(function (res) {
                    debugger;
                    $scope.NewItemTitle = "";
                    $scope.Items.push(res);
                });
            }
        }
    }]);