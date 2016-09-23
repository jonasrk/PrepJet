function redirectHome() {
    window.location = "mac_start.html";
}

(function () {
    'use strict';
    //var trim_array = [];
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            app.initialize();
            jQuery('#trim_space').click(trimSpaces);
            jQuery('#homeButton').click(redirectHome);
            document.getElementById('#test').onclick = function () {
                console.log("print test");
            }

            document.getElementById("resultClose").onclick = function () {
                document.getElementById('resultDialog').style.visibility = 'hidden';
                window.location = "harmonize.html";
            }
            document.getElementById("resultOk").onclick = function () {
                document.getElementById('resultDialog').style.visibility = 'hidden';
                window.location = "harmonize.html";
            }

        });
    };





})();