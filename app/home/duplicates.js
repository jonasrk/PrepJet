function fuzzyPro() {
    document.getElementById('showEnterprise').style.visibility = 'visible';
    document.getElementById('fuzzymatch').checked = false;
}

function redirectHome() {
    window.location = "mac_start.html";
}

(function () {
    // 'use strict';
    var sorted_rows = [];
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            Office.context.document.settings.set('same_header_duplicates', false);
            Office.context.document.settings.set('last_clicked_function', "duplicates.html");
            if (Office.context.document.settings.get('prepjet_loaded_before') == null) {
                Office.context.document.settings.set('backup_sheet_count', 1);
                Office.context.document.settings.set('prepjet_loaded_before', true);
                Office.context.document.settings.saveAsync();
                window.location = "intro.html";
            }

            app.initialize();
            populateCheckboxes();


            $('#bt_detect_duplicates').click(detectDuplicates);
            $('#checkbox_all').click(checkCheckbox);
            $('#buttonOk').click(highlightHeader);
            $('#homeButton').click(redirectHome);


            //show and hide error message when columns have same header name
            document.getElementById("buttonClose").onclick = function () {
                document.getElementById('showEmbeddedDialog').style.visibility = 'hidden';
            }

            //show and hide message about PrepJet Pro when hovering over fuzzy matching
            document.getElementById('buttonCloseEnterprise').onclick = function () {
                document.getElementById('showEnterprise').style.visibility = 'hidden';
            }
            document.getElementById('buttonOkEnterprise').onclick = function () {
                document.getElementById('showEnterprise').style.visibility = 'hidden';
            }

            //show and hide help callout
            document.getElementById("help_icon").onclick = function () {
                document.getElementById('helpCallout').style.visibility = 'visible';
            }
            document.getElementById("closeCallout").onclick = function () {
                document.getElementById('helpCallout').style.visibility = 'hidden';
            }

            document.getElementById("refresh_icon").onclick = function () {
                window.location = "duplicates.html";
            }

            //hide result message
            document.getElementById("resultClose").onclick = function () {
                document.getElementById('resultDialog').style.visibility = 'hidden';
                window.location = "duplicates.html";
            }
            document.getElementById("resultOk").onclick = function () {
                document.getElementById('resultDialog').style.visibility = 'hidden';
                window.location = "duplicates.html";
            }


            /*Excel.run(function (ctx) {

                var myBindings = Office.context.document.bindings;
                var worksheetname = ctx.workbook.worksheets.getActiveWorksheet();

                worksheetname.load('name')

                return ctx.sync().then(function() {

                    Office.context.document.addHandlerAsync("documentSelectionChanged", myViewHandler, function(result){}
                    );

                    // Event handler function for changing the worksheet.
                    function myViewHandler(eventArgs){
                        Excel.run(function (ctx) {
                            var selectedSheet = ctx.workbook.worksheets.getActiveWorksheet();
                            selectedSheet.load('name');
                            return ctx.sync().then(function () {
                                if (selectedSheet.name != worksheetname.name) {
                                    window.location = "duplicates.html"
                                }
                            });
                        });
                    }


                    function bindFromPrompt() {

                        var myBindings = Office.context.document.bindings;
                        var name_worksheet = worksheetname.name;
                        var myAddress = name_worksheet.concat("!1:1");

                        myBindings.addFromNamedItemAsync(myAddress, "matrix", {id:'myBinding'}, function (asyncResult) {
                            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                                write('Action failed. Error: ' + asyncResult.error.message);
                            } else {
                                write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);

                                function addHandler() {
                                    Office.select("bindings#myBinding").addHandlerAsync(
                                        Office.EventType.BindingDataChanged, dataChanged);
                                }

                                addHandler();
                                displayAllBindings();

                            }
                        });
                    }

                bindFromPrompt();

                function displayAllBindings() {
                    Office.context.document.bindings.getAllAsync(function (asyncResult) {
                        var bindingString = '';
                        for (var i in asyncResult.value) {
                            bindingString += asyncResult.value[i].id + '\n';
                        }
                    });
                }

                function dataChanged(eventArgs) {
                    window.location = "duplicates.html";
                }

                // Function that writes to a div with id='message' on the page.
                function write(message){
                    console.log(message);
                }

                });
            }).catch(function(error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });*/

        });
    };


    function highlightHeader() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange(true);
            var firstCell = range.getColumn(0);
            var firstCol = firstCell.getEntireColumn();
            var tmpRow = range.getRow(0);
            var firstRow = tmpRow.getEntireRow();


            //get used range in active Sheet
            range.load('text');
            firstRow.load('address');
            firstCol.load('address');
            worksheet.load('name');

            return ctx.sync().then(function() {

                var tmp_offset = firstCol.address;
                var col_offset = tmp_offset.substring(tmp_offset.indexOf("!") + 1, tmp_offset.indexOf(":"));
                var tmp_row = firstRow.address;
                var row_offset = Number(tmp_row.substring(tmp_row.indexOf("!") + 1, tmp_row.indexOf(":")));
                var add_col = getNumberFromChar(col_offset);

                for (var run = 0; run < range.text[0].length - 1; run++) {
                    for (var run2 = run + 1; run2 < range.text[0].length; run2++) {
                        if (range.text[0][run] == range.text[0][run2] && range.text[0][run] != "") {
                            document.getElementById('showEmbeddedDialog').style.visibility = 'hidden';
                            highlightContentNew(worksheet.name, getCharFromNumber(run + add_col) + row_offset, '#EA7F04', function () {});
                            highlightContentNew(worksheet.name, getCharFromNumber(run2 + add_col) + row, '#EA7F04', function () {});
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


    function populateCheckboxes() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange(true);
            var firstCell = range.getColumn(0);
            var firstCol = firstCell.getEntireColumn();
            var tmpRow = range.getRow(0);
            var firstRow = tmpRow.getEntireRow();

            range.load('address');
            range.load('text');
            firstRow.load('address');
            firstCol.load('address');

            return ctx.sync().then(function() {

                var tmp_offset = firstCol.address;
                var col_offset = tmp_offset.substring(tmp_offset.indexOf("!") + 1, tmp_offset.indexOf(":"));
                var tmp_row = firstRow.address;
                var row_offset = Number(tmp_row.substring(tmp_row.indexOf("!") + 1, tmp_row.indexOf(":")));
                var add_col = getNumberFromChar(col_offset);

                for (var run = 0; run < range.text[0].length - 1; run++) {
                    for (var run2 = run + 1; run2 < range.text[0].length; run2++) {
                        if (range.text[0][run] == range.text[0][run2] && range.text[0][run] != "") {
                            document.getElementById('showEmbeddedDialog').style.visibility = 'visible';
                            Office.context.document.settings.set('same_header_duplicates', true);
                        }
                    }
                }

                for (var i = 0; i < range.text[0].length; i++) { // .text[0] is the first row of a range
                    if (range.text[0][i] != ""){
                        addNewCheckboxToContainer (range.text[0][i], "duplicates_column_checkbox" ,"checkboxes_duplicates");
                    }
                    else {
                        var colchar = getCharFromNumber(i + add_col);
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
            var range = range_all.getUsedRange(true);
            var firstCell = range.getColumn(0);
            var firstCol = firstCell.getEntireColumn();
            var tmpRow = range.getRow(0);
            var firstRow = tmpRow.getEntireRow();

            range.load('text');
            firstRow.load('address');
            firstCol.load('address');

            return ctx.sync().then(function() {

                var tmp_offset = firstCol.address;
                var col_offset = tmp_offset.substring(tmp_offset.indexOf("!") + 1, tmp_offset.indexOf(":"));
                var tmp_row = firstRow.address;
                var row_offset = Number(tmp_row.substring(tmp_row.indexOf("!") + 1, tmp_row.indexOf(":")));
                var add_col = getNumberFromChar(col_offset);

                if (document.getElementById('checkbox_all').checked == true) {
                    for (var i = 0; i < range.text[0].length; i++) {
                        if (range.text[0][i] != "") {
                            document.getElementById(range.text[0][i]).checked = true;
                        }
                        else {
                            document.getElementById("Column " + getCharFromNumber(i + add_col)).checked = true;
                        }
                    }
                }
                else {
                    for (var i = 0; i < range.text[0].length; i++) {
                        if (range.text[0][i] != "") {
                            document.getElementById(range.text[0][i]).checked = false;
                        }
                        else {
                            document.getElementById("Column " + getCharFromNumber(i + add_col)).checked = false;
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
            var range = range_all.getUsedRange(true);
            var firstCol = range.getRow(1);
            var lastCol = range.getLastColumn();

            var firstCell = range.getColumn(0);
            var firstColumn = firstCell.getEntireColumn();
            var endColor = lastCol.getEntireColumn();
            var tmpRow = range.getRow(0);
            var firstRow = tmpRow.getEntireRow();

            range.load('address');
            range.load('text');
            worksheet.load('name');
            firstCol.load('address');
            lastCol.load('address');
            firstRow.load('address');
            firstColumn.load('address');
            endColor.load('address');

            return ctx.sync().then(function() {

                var tmp_offset = firstColumn.address;
                var col_offset = tmp_offset.substring(tmp_offset.indexOf("!") + 1, tmp_offset.indexOf(":"));
                var tmp_row = firstRow.address;
                var row_offset = Number(tmp_row.substring(tmp_row.indexOf("!") + 1, tmp_row.indexOf(":")));
                var add_col = getNumberFromChar(col_offset);
                var startCell = col_offset + row_offset;

                backupForUndo(range, startCell, add_col, row_offset);

                var columns_to_check = [];

                for (var k = 0; k < range.text[0].length; k++) { // .text[0] is the first row of a range
                    for (var l = 0; l < checked_checkboxes.length; l++) { // TODO throws error if none are checked
                        if (checked_checkboxes[l].id == range.text[0][k] || checked_checkboxes[l].id == "Column " + getCharFromNumber(k + add_col)) {
                            columns_to_check.push(k);
                        }
                    }
                }

                var strings_to_sort  = [];

                for (var i = 1; i < range.text.length; i++) {
                    var this_row = [];
                    for (var j = 0; j < columns_to_check.length; j++) {
                        var row_number = i + row_offset;
                        this_row.push([range.text[i][columns_to_check[j]], getCharFromNumber(columns_to_check[j] + add_col) + row_number, row_number]);
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

                var dup_count = 1;
                var dup_index = 1;
                for (var o = 1; o < strings_to_sort.length; o++){
                    if (arraysEqual(strings_to_sort[o], strings_to_sort[o - 1])){
                        if (dup_count == 1) {
                            duplicates.push([strings_to_sort[o], dup_index]);
                            duplicates.push([strings_to_sort[o - 1], dup_index]);
                            dup_count += 1;
                        }
                        else {
                            duplicates.push([strings_to_sort[o], dup_index]);
                        }
                    }
                    else {
                        dup_count = 1;
                        dup_index += 1;
                    }
                }

                function colorDup(duplicates_input, int) {
                    var color = "#EA7F04";
                    for (var m = 0; m < duplicates_input.length; m++){
                        if (m > 0 && duplicates_input[m][1] != duplicates_input[m-1][1]){
                            color = getRandomColor();
                        }

                        for (var n = 0; n < duplicates_input[m].length; n++){
                            for (var o = 0; o < duplicates_input[m][n].length; o++) {
                                highlightContentNew(worksheet.name, duplicates_input[m][n][o][int], color, function () {});
                            }

                        }
                    }
                }

                if(document.getElementById('duplicatesort').checked == false) {
                    colorDup(duplicates, 1);
                } else {

                    var dup_length = duplicates.length;
                    var sheet_row = row_offset + 1;

                    var row_numbers = [];
                    for (var run = 0; run < duplicates.length; run++) {
                        row_numbers.push(duplicates[run][0][0][2]);
                    }

                    var text = [];
                    var color_check = []
                    var data_index = 0;
                    for (var run = 0; run < row_numbers.length; run++) {
                        data_index = row_numbers[run];
                        text.push(range.text[data_index - row_offset]);
                        color_check.push(duplicates[run][1]);
                    }

                    for (var run = 1; run < range.text.length; run ++) {
                        var check = 0;
                        for (var k = 0; k < row_numbers.length; k++) {
                            if((run + row_offset) == row_numbers[k]) {
                                check = 1;
                            }
                        }
                        if (check == 0) {
                            text.push(range.text[run]);
                            console.log(range.text[run])
                        }
                    }

                    var color = '#EA7F04';

                    var start_col = firstCol.address.substring(firstCol.address.indexOf("!") + 1, firstCol.address.indexOf(":"));
                    var end_col = lastCol.address.substring(lastCol.address.indexOf(":") + 1);

                    addContentNew(worksheet.name, start_col + ":" + end_col, text, function () {});

                    var end_color = endColor.address.substring(endColor.address.indexOf("!") + 1, endColor.address.indexOf(":"));
                    for (var row = 0; row < text.length; row++) {
                        if (sheet_row < (row_numbers.length + row_offset + 2)) {
                            if (row > 0 && color_check[row] == color_check[row - 1]) {
                                var insert_address = col_offset + sheet_row + ":" +  end_color + sheet_row;
                                highlightContentNew(worksheet.name, insert_address, color, function () {});
                            }
                        }
                        sheet_row += 1;
                    }

                }


                if (document.getElementById('createBackup').checked == true) {
                    var sheet_count = Office.context.document.settings.get('backup_sheet_count') + 1;
                    Office.context.document.settings.set('backup_sheet_count', sheet_count);
                    Office.context.document.settings.saveAsync();
                    var newName = worksheet.name + "(" + sheet_count + ")";
                    addBackupSheet(newName, startCell, add_col, row_offset, function() {
                        var txt = document.createElement("p");
                        txt.className = "ms-font-xs ms-embedded-dialog__content__text";
                        txt.innerHTML = "PrepJet found " + duplicates.length + " duplicate rows."
                        document.getElementById('resultText').appendChild(txt);

                        document.getElementById('resultDialog').style.visibility = 'visible';
                    });

                    }
                    else {
                        var txt = document.createElement("p");
                        txt.className = "ms-font-xs ms-embedded-dialog__content__text";
                        txt.innerHTML = "PrepJet found " + duplicates.length + " duplicate rows."
                        document.getElementById('resultText').appendChild(txt);

                        document.getElementById('resultDialog').style.visibility = 'visible';
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