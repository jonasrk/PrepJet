(function () {
    'use strict';
    var sorted_rows = [];
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            Office.context.document.settings.set('last_clicked_function', "duplicates.html");
            if (Office.context.document.settings.get('prepjet_loaded_before') == null) {
                Office.context.document.settings.set('prepjet_loaded_before', true);
                Office.context.document.settings.saveAsync();
                window.location = "intro.html";
            }

            app.initialize();
            populateCheckboxes();

            $('#bt_detect_duplicates').click(detectDuplicates);
            $('#checkbox_all').click(checkCheckbox);

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
                    if (range.text[0][i] != ""){
                        addNewCheckboxToContainer (range.text[0][i], "duplicates_column_checkbox" ,"checkboxes_duplicates");
                    }
                    else {
                        var colchar = getCharFromNumber(i);
                        addNewCheckboxToContainer ("Column " + colchar, "duplicates_column_checkbox" ,"checkboxes_duplicates");
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


    function checkCheckbox() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('text');

            return ctx.sync().then(function() {
                if (document.getElementById('checkbox_all').checked == true) {
                    for (var i = 0; i < range.text[0].length; i++) {
                        if (range.text[0][i] != "") {
                            document.getElementById(range.text[0][i]).checked = true;
                        }
                        else {
                            document.getElementById("Column " + getCharFromNumber(i)).checked = true;
                        }
                    }
                }
                else {
                    for (var i = 0; i < range.text[0].length; i++) {
                        if (range.text[0][i] != "") {
                            document.getElementById(range.text[0][i]).checked = false;
                        }
                        else {
                            document.getElementById("Column " + getCharFromNumber(i)).checked = false;
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
                        if (checked_checkboxes[l].id == range.text[0][k] || checked_checkboxes[l].id == "Column " + getCharFromNumber(k)) {
                            columns_to_check.push(k);
                        }
                    }
                }

                var strings_to_sort  = [];


                for (var i = 1; i < range.text.length; i++) {
                    var this_row = [];
                    for (var j = 0; j < columns_to_check.length; j++) {
                        var row_number = i + 1;
                        this_row.push([range.text[i][columns_to_check[j]], getCharFromNumber(columns_to_check[j]) + row_number]);
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

                function colorDup(duplicates_input, int) {
                    var color = "#EA7F04";
                    for (var m = 0; m < duplicates_input.length; m++){
                        if (m > 0 && !arraysEqual(duplicates_input[m], duplicates_input[m-1])){
                            // generate new random color
                            color = getRandomColor();
                        }

                        for (var n = 0; n < duplicates_input[m].length; n++){
                            highlightContentInWorksheet(worksheet, duplicates_input[m][n][int], color);
                        }
                    }
                }

                if(document.getElementById('duplicatesort').checked == false) {
                    colorDup(duplicates, 1);
                    window.location = "duplicates.html";
                }
                else {

                    var dup_length = duplicates.length;
                    var sheet_row = 2;
                    var row_array = [];

                    for (var run = 0; run < dup_length; run++) {
                        row_array[run] = duplicates[run][0][1];
                    }

                    var row_numbers = [];
                    for (var run = 0; run < row_array.length; run++) {
                        row_numbers[run] = Number(row_array[run].substring(1));
                    }

                    var text = [];
                    var data_index = 0;
                    for (var run = 0; run < row_numbers.length; run++) {
                        data_index = row_numbers[run] - 1;
                        text.push(range.text[data_index]);
                    }

                    for (var run = 1; run < range.text.length; run ++) {
                        var check = 0;
                        for (var k = 0; k < row_numbers.length; k++) {
                            if((run + 1) == row_numbers[k]) {
                                check = 1;
                            }
                        }
                        if (check == 0) {
                            text.push(range.text[run]);
                        }
                    }

                    var color = '#EA7F04';
                    for (var row = 0; row < text.length; row++) {
                        for(var col = 0; col < range.text[0].length; col++) {
                            var columnchar = getCharFromNumber(col)
                            addContentToWorksheet(worksheet, columnchar + sheet_row, text[row][col])
                            if (sheet_row < (row_numbers.length + 2)) {
                                highlightContentInWorksheet(worksheet, columnchar + sheet_row ,color)
                            }
                        }
                        sheet_row += 1;
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