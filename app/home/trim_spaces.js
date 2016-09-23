function redirectHome() {
    window.location = "mac_start.html";
}

(function () {
    'use strict';
    //var trim_array = [];
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            jQuery('#test').click(printTest);


        });
    };

    function printTest() {
        console.log("testtest");
    }

})();