(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            Office.context.document.settings.set('last_clicked_function', "outlier.html");
            if (Office.context.document.settings.get('prepjet_loaded_before') == null) {
                Office.context.document.settings.set('prepjet_loaded_before', true);
                Office.context.document.settings.saveAsync();
                window.location = "intro.html";
            }

            app.initialize();

            populateColumnDropdown();

            $('#bt_detect_outliers').click(detectOutliers);

        });
    };


    function populateColumnDropdown() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('address');
            range.load('text');
            return ctx.sync().then(function() {
                for (var i = 0; i < range.text[0].length; i++) {

                    var el = document.createElement("option");
                    el.value = range.text[0][i];
                    el.textContent = range.text[0][i];
                    document.getElementById("outlier_column_dropdown").appendChild(el);

                }

                $(".outlier_column_dropdown_container").Dropdown();
            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

    }

    function detectOutliers() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('address');
            range.load('text');
            return ctx.sync().then(function() {

                var selected_column = document.getElementById('outlier_column_dropdown').value; // TODO better reference by ID than name

                // iterate over columns

                for (var k = 0; k < range.text[0].length; k++){
                    if (selected_column == range.text[0][k]){

                        var values = [];

                        for (var j = 1; j < range.text.length; j++){

                            values.push(range.text[j][k]);

                        }

                        // High outliers are anything beyond the 3rd quartile + 1.5 * the inter-quartile range (IQR)
                        // Low outliers are anything beneath the 1st quartile - 1.5 * IQR

                        values.sort();

                        var Q1, Q2, Q3 = 0;

                        var q1Arr = (values.length % 2 == 0) ? values.slice(0, (values.length / 2)) : values.slice(0, Math.floor(values.length / 2));
                        var q2Arr =  values;
                        var q3Arr = (values.length % 2 == 0) ? values.slice((values.length / 2), values.length) : values.slice(Math.ceil(values.length / 2), values.length);
                        Q1 = medianX(q1Arr);
                        Q2 = medianX(q2Arr);
                        Q3 = medianX(q3Arr);
                        function medianX(medianArr) {
                            var count = medianArr.length;
                            var median = (count % 2 == 0) ? (medianArr[(medianArr.length/2) - 1] + medianArr[(medianArr.length / 2)]) / 2:medianArr[Math.floor(medianArr.length / 2)];
                            return median;
                        }

                        var iqr = Q3 - Q1;

                        var thrsh_low = Q1 - (1.5 * iqr);
                        var thrsh_high = Q3 + (1.5 * iqr); // TODO do not hardcode

                        for (var j = 1; j < range.text.length; j++){

                            var sheet_row = j + 1;
                            var address = getCharFromNumber(k + 1) + sheet_row;

                            if (range.text[j][k] < thrsh_low){
                                highlightContentInWorksheet(worksheet, address, "red");
                            } else if (range.text[j][k] > thrsh_high){
                                highlightContentInWorksheet(worksheet, address, "red");
                            }

                        }

                    }
                }



            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

})();