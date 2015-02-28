'use strict';

SPKnight.TestApp.App.controller('ListOpsController', ['$scope', '$rootScope', '$filter', 'AngularSPREST', 'AngularSPCSOM',
    function ($scope, $rootScope, $filter, $angularSPRest, $angularSPCSOM) {
        $scope.Items = [];
        $scope.Method = "REST";

        $scope.GetListItems = function GetListItems()
        {
            if($scope.Method === "REST")
            {
                $angularSPRest.GetListItems("TestList", "/TestAngularSP").then(function (res) {
                    $scope.Items = res.data.d.results;
                });
            }
            else
            {
                var res = $angularSPCSOM.GetListItems("TestList", "/TestAngularSP");
                res.Promise.then(function () {
                    debugger;
                });
            }
        }
        $scope.UpdateItem = function UpdateItem(item)
        {
            if ($scope.Method === "REST") {
                $angularSPRest.UpdateListItem(item.ID, "TestList", "/TestAngularSP", { Title: item.Title }).then(function (res) {
                    debugger;
                });
            }
            else {
                $angularSPCSOM.GetListItems("TestList", "/TestAngularSP").then(function (res) {
                    debugger;
                });
            }
        }
        $scope.DeleteItem = function DeleteItem(item)
        {
            if ($scope.Method === "REST") {
                $angularSPRest.DeleteListItem(item.ID, "TestList", "/TestAngularSP").then(function (res) {
                    var index = $scope.Items.indexOf(item);
                    $scope.Items.splice(index, 1);
                });
            }
            else {
                $angularSPCSOM.GetListItems("TestList", "/TestAngularSP").then(function (res) {
                    debugger;
                });
            }
        }
    }]);