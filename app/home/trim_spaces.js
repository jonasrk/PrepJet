function resultOK() {
    document.getElementById('resultDialog').style.visibility = 'hidden';
    window.location = "harmonize.html";
}

function resultClose() {
    document.getElementById('resultDialog').style.visibility = 'hidden';
    window.location = "harmonize.html";
}

function redirectHome() {
    window.location = "mac_start.html";
}

(function(){
    //'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function(reason){
        jQuery(document).ready(function(){
            app.initialize();

            jQuery('#trim_space').click(getDataFromSelection);
            //jQuery('#trim_space').click(testWriting);

            jQuery('#resultOk').click(resultOK);
            jQuery('#resultOk').click(resultClose);
            jQuery('#homeButton').click(redirectHome);


        });
    };

    function testWriting() {
        var test = [["eins"], ["zwei"]];
        Office.context.document.setSelectedDataAsync(test, { coercionType: Office.CoercionType.Matrix }, function(result) {
            console.log(result);
        });
    }
    // Reads data from current document selection and displays a notification
    function getDataFromSelection(){
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Matrix,
            function(result){
                getSelectedData(function(result){

                    if (result != null) {
                        var countTrim = 0;
                        var trim_array = [result.map(function (item) {
                            return item.map(function (item) {
                                if (item) {
                                    var newitem = item.trim();
                                    if (item != newitem) {
                                        countTrim++;
                                    }

                                    return newitem;
                                }
                            });
                        })];
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

                });
        });
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
