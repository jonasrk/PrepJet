(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {
            app.initialize();

            populateCheckboxes();

            $('#bt_detect_duplicates').click(detectDuplicates);

        });
    };


    function populateCheckboxes() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('address');
            range.load('text');
            return ctx.sync().then(function() {
                for (var i = 0; i < range.text[0].length; i++) { // .text[0] is the first row of a range
                    addNewCheckboxToContainer (range.text[0][i], "duplicates_column_checkbox" ,"checkboxes_duplicates");
                }
            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

    }

    function detectDuplicates() {

        var checked_checkboxes = getCheckedBoxes("duplicates_column_checkbox");

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('address');
            range.load('text');
            return ctx.sync().then(function() {

                var columns_to_check = [];

                for (var k = 0; k < range.text[0].length; k++) { // .text[0] is the first row of a range

                    for (var l = 0; l < checked_checkboxes.length; l++) { // TODO throws error if none are checked

                        if (checked_checkboxes[l].id == range.text[0][k]) {

                            columns_to_check.push(k);

                        }

                    }

                }

                var strings_to_sort  = [];


                for (var i = 1; i < range.text.length; i++) {

                    var this_row = [];

                    for (var j = 0; j < columns_to_check.length; j++) {
                        var row_number = i + 1;
                        this_row.push([range.text[i][columns_to_check[j]], getCharFromNumber(columns_to_check[j] + 1) + row_number]);

                    }

                    strings_to_sort.push(this_row);

                }

                function Comparator(a, b) {

                    for (var i = 0; i < checked_checkboxes.length; i++){
                        if (a[i][0] < b[i][0]) return -1;
                        if (a[i][0] > b[i][0]) return 1;
                    }
                    return 0;
                }

                strings_to_sort.sort(Comparator);

                var duplicates = [];

                function arraysEqual(a, b) {
                    if (a === b) return true;
                    if (a == null || b == null) return false;
                    if (a.length != b.length) return false;

                    // If you don't care about the order of the elements inside
                    // the array, you should sort both arrays here.

                    for (var i = 0; i < a.length; ++i) {
                        if (a[i][0] !== b[i][0]) return false;
                    }
                    return true;
                }

                for (var o = 1; o < strings_to_sort.length; o++){
                    if (arraysEqual(strings_to_sort[o] ,strings_to_sort[o - 1])){
                        duplicates.push(strings_to_sort[o]);
                        duplicates.push(strings_to_sort[o - 1]);
                    }

                }

                for (var m = 0; m < duplicates.length; m++){

                    for (var n = 0; n < duplicates[m].length; n++){

                        highlightContentInWorksheet(worksheet, duplicates[m][n][1]);

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