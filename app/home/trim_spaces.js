function redirectHome() {
    window.location = "mac_start.html";
}

(function () {
    // 'use strict';
    var trim_array = [];
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            app.initialize();
            jQuery('#trim_space').click(trimSpaces);
            jQuery('#homeButton').click(redirectHome);

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


    // Reads data from current document selection and displays a notification
    function trimSpaces(){

           getSelectedData(function(result){

                if (result != null) {
                    var countTrim = 0;
                    var trim_array = result.map(function (item) {
                        return item.map(function (item) {
                            if (item) {
                                var newitem = item.trim();
                                if (item != newitem) {
                                    countTrim++;
                                }
                                return newitem;
                            }
                        });
                    });
                }
                Office.context.document.setSelectedDataAsync(trim_array, { valueFormat: Office.ValueFormat.Formatted }, function(result){
                        if (result.status == "succeeded") {
                            var txt = document.createElement("p");
                            txt.className = "ms-font-xs ms-embedded-dialog__content__text";
                            txt.innerHTML = "PrepJet trimed " + countTrim + " spaces in the selected range."
                            document.getElementById('resultText').appendChild(txt);
                            document.getElementById('resultDialog').style.visibility = 'visible';
                        } else {
                            console.log("An error occured. Please select a range and try again.");
                        }
                    });
            }
        );

    }

    function getSelectedData(callback) {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Matrix, { valueFormat: Office.ValueFormat.Formatted },
        function (result) {
            if (result.status == "succeeded") {
                callback(result.value);
            }
            else {
                console.log("error");
            }
        });
    }


})();