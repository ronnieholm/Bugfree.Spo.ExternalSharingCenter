(function () {
    'use strict';

    angular.element(document).ready(function () {
        angular.bootstrap(document, ['app']);
    });

    var app = angular.module('app', ['ui.date']);
})();