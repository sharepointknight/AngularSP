'use strict';

SPKnight.TestApp.App.controller('SearchOpsController', ['$scope', '$rootScope', '$filter', 'AngularSPREST', 'AngularSPCSOM',
    function ($scope, $rootScope, $filter, $angularSPRest, $angularSPCSOM) {
        $scope.Items = [];
        $scope.Method = "REST";

        $scope.GetListItems = function GetListItems()
        {
            if($scope.Method === "REST")
            {
                $angularSPRest.GetListItems("TestList", "/TestAngularSP").then(function (res) {
                    debugger;
                });
            }
            else
            {
                $angularSPCSOM.GetListItems("TestList", "/TestAngularSP").then(function (res) {
                    debugger;
                });
            }
        }
    }]);