function redirectHome() {
    window.location = "mac_start.html";
}

(function () {
    'use strict';
    //var trim_array = [];
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            document.getElementById('#test').onclick = function () {
                console.log("print test");
            }


        });
    };



})();