(function(){
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function(reason){
        jQuery(document).ready(function(){
            app.initialize();

            jQuery('#test').click(getDataFromSelection);
            //jQuery('#replace-checked-values').click(replaceCheckedValues);
        });
    };

    // Reads data from current document selection and displays a notification
    function getDataFromSelection(){
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function(result){
                getSelectedData(function(result){

                    if (result != null) {
                        var countTrim = 0;
                        var trim_array = result.map(function (item) {
                            return item.map(function (item) {
                                if (item) {
                                    console.log("third");
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
                            var txt = document.createElement("label");
                            txt.innerHTML = "success";
                            document.getElementById('explanation').appendChild(txt);
                        } else {
                            var txt = document.createElement("label");
                            txt.innerHTML = "not succeeded";
                            document.getElementById('explanation').appendChild(txt);
                            console.log("An error occured. Please select a range and try again.");
                        }
                    });

                    var txt = document.createElement("label");
                    txt.innerHTML = "testtestetst";
                    document.getElementById('explanation').appendChild(txt);

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
