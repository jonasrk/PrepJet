function resultOK() {
    document.getElementById('resultDialog').style.visibility = 'hidden';
    window.location = "mac_start.html";
}

function resultClose() {
    document.getElementById('resultDialog').style.visibility = 'hidden';
    window.location = "mac_start.html";
}

function redirectHome() {
    window.location = "mac_start.html";
}

(function(){
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function(reason){
        jQuery(document).ready(function(){
            app.initialize();

            jQuery('#inconsistency').click(getDataFromSelection);

            jQuery('#resultOk').click(resultOK);
            jQuery('#resultOk').click(resultClose);
            jQuery('#homeButton').click(redirectHome);


        });
    };

    function getDataType(item) {
        var datatype = "string";
        return datatype;
    }

    // Reads data from current document selection and displays a notification
    function getDataFromSelection(){
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function(result){
                getSelectedData(function(result){

                    if (result != null) {
                        var countIncon = 0;
                        var type_array = result.map(function (item) {
                            return item.map(function (item) {
                                if (item) {
                                    var newitem = getDataType(item);
                                    if (item != newitem) {
                                        countIncon++;
                                    }
                                    return newitem;
                                }
                            });
                        });
                    }

                    Office.context.document.setSelectedDataAsync(type_array, { valueFormat: Office.ValueFormat.Formatted }, function(result){
                        if (result.status == "succeeded") {
                            var txt = document.createElement("p");
                            txt.className = "ms-font-xs ms-embedded-dialog__content__text";
                            txt.innerHTML = "PrepJet found " + countIncon + " data entries with inconsitent data type."
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
