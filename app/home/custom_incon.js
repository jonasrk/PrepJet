function displayBetween(){
    if(document.getElementById('charOptions').value == "between" || document.getElementById('charOptions').value == "notbetween") {
        $('#between').show();
    }
    else {
        $('#between').hide();
    }
}

function redirectHome() {
    window.location = "mac_start.html";
}


(function () {
    // 'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            Office.context.document.settings.set('same_header_trim', false);
            Office.context.document.settings.set('last_clicked_function', "custom_incon.html");
            if (Office.context.document.settings.get('prepjet_loaded_before') == null) {
                Office.context.document.settings.set('prepjet_loaded_before', true);
                Office.context.document.settings.saveAsync();
                window.location = "intro.html";
            }

            app.initialize();
            fillColumn();

            $('#helpCallout').hide();
            $('#between').hide();
            $(".dropdown_table").Dropdown();
            $('#custom_incon').click(screenIncon);
            $('#buttonOk').click(highlightHeader);
            $('#homeButton').click(redirectHome);

            $zopim(function() {
                $zopim.livechat.window.hide();
            });

            //Show and hide error message if columns have same header name
            document.getElementById("buttonClose").onclick = function () {
                document.getElementById('showEmbeddedDialog').style.visibility = 'hidden';
            }

            document.getElementById("refresh_icon").onclick = function () {
                window.location = "custom_incon.html";
            }

            //hide result message
            document.getElementById("resultClose").onclick = function () {
                document.getElementById('resultDialog').style.visibility = 'hidden';
                window.location = "custom_incon.html";
            }
            document.getElementById("resultOk").onclick = function () {
                document.getElementById('resultDialog').style.visibility = 'hidden';
                window.location = "custom_incon.html";
            }


            /*Excel.run(function (ctx) {

                var myBindings = Office.context.document.bindings;
                var worksheetname = ctx.workbook.worksheets.getActiveWorksheet();

                var headRange_all = worksheetname.getRange();
                var headRange = headRange_all.getUsedRange(true);

                worksheetname.load('name')
                headRange.load('text');

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
                                    window.location = "trim_spaces.html"
                                }
                            });
                        });
                    }

                    //function to check whether header entries are changed
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
                    window.location = "trim_spaces.html";
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
                            highlightContentNew(worksheet.name, getCharFromNumber(run2 + add_col) + row_offset, '#EA7F04', function () {});
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


    function fillColumn(){

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

                for (var run = 0; run < range.text[0].length - 1; run++) {
                    for (var run2 = run + 1; run2 < range.text[0].length; run2++) {
                        if (range.text[0][run] == range.text[0][run2] && range.text[0][run] != "") {
                            document.getElementById('showEmbeddedDialog').style.visibility = 'visible';
                            Office.context.document.settings.set('same_header_split', true);
                        }
                    }
                }

                for (var i = 0; i < range.text[0].length; i++) {
                    var el = document.createElement("option");
                    if (range.text[0][i] != "") {
                        el.value = range.text[0][i];
                        el.textContent = range.text[0][i];
                    }
                    else {
                        el.value = "Column " + getCharFromNumber(i + add_col);
                        el.textContent = "Column " + getCharFromNumber(i + add_col);
                    }
                    document.getElementById("column_options").appendChild(el);
                }
                $(".dropdown_table_col").Dropdown();
                $("span.ms-Dropdown-title:empty").text(range.text[0][0]);
            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

    }


    function screenIncon() {

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
            worksheet.load('name');
            firstRow.load('address');
            firstCol.load('address');

            return ctx.sync().then(function() {

                var tmp_offset = firstCol.address;
                var col_offset = tmp_offset.substring(tmp_offset.indexOf("!") + 1, tmp_offset.indexOf(":"));
                var tmp_row = firstRow.address;
                var row_offset = Number(tmp_row.substring(tmp_row.indexOf("!") + 1, tmp_row.indexOf(":")));
                var add_col = getNumberFromChar(col_offset);
                var startCell = col_offset + row_offset;

                backupForUndo(range, startCell, add_col, row_offset);

                var selected_identifier = document.getElementById('column_options').value;
                var charCount = Number(document.getElementById('charCountInput').value);
                var charIncluded = document.getElementById('includeChar').value;
                var charNotIncluded = document.getElementById('notIncludeChar').value;
                var charOperator = document.getElementById('charOptions').value;

                backupForUndo(range);

                var header = 0;
                for (var k = 0; k < range.text[0].length; k++){
                    if (selected_identifier == range.text[0][k] || selected_identifier == "Column " + getCharFromNumber(k + add_col)){
                        header = k;
                    }
                }

                var countIncons = 0;
                var color = "#EA7F04";

                if (charCount != 0) {
                    for (var k = 1; k < range.text.length; k++) {
                        var string_length = range.text[k][header].length;
                        checkCharCount(string_length, range.text[k][header]);
                    }
                }

                function checkCharCount(str_length, input_str) {
                    var check_cond = 0;
                    if (charOperator == "equal") {
                        if (str_length != charCount) {
                            check_cond = 1;
                        }
                    }
                    if (charOperator == "smaller") {
                        if (str_length >= charCount) {
                            check_cond = 1;
                        }
                    }
                    if (charOperator == "greater") {
                        if (str_length <= charCount) {
                            check_cond = 1;
                        }
                    }
                    if (charOperator == "inequal") {
                        if (str_length == charCount) {
                            check_cond = 1;
                        }
                    }
                    if (charOperator == "between") {
                        var upperRange = Number(document.getElementById('betweenInput').value);
                        if (str_length < charCount || str_length > upperRange) {
                            check_cond = 1;
                        }
                    }
                    if (charOperator == "notbetween") {
                        var upperRange = Number(document.getElementById('betweenInput').value);
                        if (str_length >= charCount && str_length <= upperRange) {
                            check_cond = 1;
                        }
                    }
                    if (check_cond == 1) {
                        countIncons += 1;
                        highlightCellNew(worksheet.name, getCharFromNumber(header + add_col) + (k + row_offset), color, function () {});
                    }

                }

                if (charIncluded != "") {
                    for (var k = 1; k < range.text.length; k++) {
                        var include_check = range.text[k][header].indexOf(charIncluded);
                        if (include_check < 0) {
                            countIncons += 1;
                            highlightContentInWorksheet(worksheet, getCharFromNumber(header + add_col) + (k + row_offset), '#EA7F04');
                            //highlightCellNew(worksheet.name, getCharFromNumber(header + add_col) + (k + row_offset), color, function () {});

                        }
                    }
                }


                if (charNotIncluded != "") {
                    for (var k = 1; k < range.text.length; k++) {
                        var notInclude_check = range.text[k][header].indexOf(charNotIncluded);
                        if (notInclude_check >= 0) {
                            countIncons += 1;
                            highlightContentInWorksheet(worksheet, getCharFromNumber(header + add_col) + (k + row_offset), '#EA7F04');
                            //highlightCellNew(worksheet.name, getCharFromNumber(header + add_col) + (k + row_offset), color, function () {});
                        }
                    }
                }

                if (document.getElementById('createBackup').checked == true) {
                    var sheet_count = Office.context.document.settings.get('backup_sheet_count') + 1;
                    Office.context.document.settings.set('backup_sheet_count', sheet_count);
                    Office.context.document.settings.saveAsync();
                    var newName = worksheet.name + "(" + sheet_count + ")";
                    addBackupSheet(newName, function() {
                        var txt = document.createElement("p");
                        txt.className = "ms-font-xs ms-embedded-dialog__content__text";
                        txt.innerHTML = "PrepJet found " + countIncons + " inconsistencies.";
                        document.getElementById('resultText').appendChild(txt);
                        document.getElementById('resultDialog').style.visibility = 'visible';
                    });

                }
                else {
                    var txt = document.createElement("p");
                    txt.className = "ms-font-xs ms-embedded-dialog__content__text";
                    txt.innerHTML = "PrepJet found " + countIncons + " inconsistencies.";
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