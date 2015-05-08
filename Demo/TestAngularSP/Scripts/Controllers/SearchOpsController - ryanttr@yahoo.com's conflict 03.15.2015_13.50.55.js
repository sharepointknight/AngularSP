'use strict';

SPKnight.TestApp.App.controller('SearchOpsController', ['$scope', '$rootScope', '$filter', 'AngularSPREST', 'AngularSPCSOM',
    function ($scope, $rootScope, $filter, $angularSPRest, $angularSPCSOM) {
        $scope.Results = null;
        $scope.Method = "REST";
        $scope.GetOrPost = "GET";
        $scope.QueryText = "";

        $scope.ExecuteSearch = function ExecuteSearch()
        {
            if($scope.Method === "REST")
            {
                var options = {
                    '__metadata': { 'type': 'Microsoft.Office.Server.Search.REST.SearchRequest' },
                    Querytext: $scope.QueryText
                }
                if ($scope.GetOrPost == "GET")
                {
                    $angularSPRest.Search.Get("/TestAngularSP", options).then(function (res) {
                        debugger;
                        $scope.Results = res;
                    });
                }
                else
                {
                    var tmp = { request: options };
                    $angularSPRest.Search.Post("/TestAngularSP", tmp).then(function (res) {
                        debugger;
                        $scope.Results = res;
                    });
                }
            }
            else
            {
                $angularSPCSOM.GetListItems("TestList", "/TestAngularSP").then(function (res) {
                    debugger;
                });
            }
        }
    }]);