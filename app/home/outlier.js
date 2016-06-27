(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {
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
                        console.log(q1Arr, q2Arr, q3Arr);
                        console.log(Q1,Q2,Q3);

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