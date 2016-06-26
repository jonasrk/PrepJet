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

                for (var l = 0; l < checked_checkboxes.length; l++){ // TODO throws error if none are checked

                    var strings_to_sort  = [];

                    for (var k = 0; k < range.text[0].length; k++) { // .text[0] is the first row of a range

                        var column_char = getCharFromNumber(1 + k);

                        if (checked_checkboxes[l].id == range.text[0][k]){

                            for (var i = 1; i < range.text.length; i++) {

                                strings_to_sort.push(range.text[i][k]);

                            }

                            strings_to_sort.sort();

                            var duplicates = [];

                            for (var o = 1; o < strings_to_sort.length; o++){

                                if (strings_to_sort[o] == strings_to_sort[o - 1]){
                                    duplicates.push(strings_to_sort[o]);

                                }

                            }

                            for (var m = 0; m < duplicates.length; m++){

                                for (var n = 1; n < range.text.length; n++) {

                                    var sheet_row = n + 1;

                                    if (duplicates[m] == range.text[n][k]){

                                        highlightContentInWorksheet(worksheet, column_char + sheet_row)

                                    }

                                }



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