'use strict';

SPKnight.TestApp.App.controller('SearchOpsController', ['$scope', '$rootScope', '$filter', 'AngularSPREST', 'AngularSPCSOM',
    function ($scope, $rootScope, $filter, $angularSPRest, $angularSPCSOM) {
        $scope.Results = null;
        $scope.Method = "REST";
        $scope.QueryText = "";

        $scope.ExecuteSearch = function ExecuteSearch()
        {
            if($scope.Method === "REST")
            {
                var options = {
                    QueryText: $scope.QueryText
                }
                $angularSPRest.Search.Get("/TestAngularSP", options).then(function (res) {
                    debugger;
                    $scope.Results = res;
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