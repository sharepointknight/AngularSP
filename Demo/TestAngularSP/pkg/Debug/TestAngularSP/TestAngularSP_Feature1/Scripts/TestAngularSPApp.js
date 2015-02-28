'use strict';

var SPKnight = SPKnight || {};
SPKnight.TestApp = SPKnight.TestApp || {};

SPKnight.TestApp.App = angular.module('TestApp', ['AngularSP', 'ngRoute']);
SPKnight.TestApp.App.filter('unsafe', function ($sce) {
    return function (val) {
        return $sce.trustAsHtml(val);
    };
});
SPKnight.TestApp.App.config(['$routeProvider',
        function($routeProvider) {
            $routeProvider.
                when('/ListOps', {
                    templateUrl: '../Content/Listops.html',
                }).
                when('/Search', {
                    templateUrl: '../Content/SearchOps.html',
                }).
                otherwise({
                    redirectTo: '/ListOps'
                });
        }]);