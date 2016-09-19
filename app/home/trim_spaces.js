function redirectHome() {
    window.location = "mac_start.html";
}

(function () {
    // 'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

        function sleepFor( sleepDuration ){
    var now = new Date().getTime();
    while(new Date().getTime() < now + sleepDuration){ /* do nothing */ }
}

sleepFor(10000);

            Office.context.document.settings.set('same_header_trim', false);
            Office.context.document.settings.set('last_clicked_function', "trim_spaces.html");
            if (Office.context.document.settings.get('prepjet_loaded_before') == null) {
                Office.context.document.settings.set('backup_sheet_count', 1);
                Office.context.document.settings.set('prepjet_loaded_before', true);
                Office.context.document.settings.saveAsync();
                window.location = "intro.html";
            }


            app.initialize();
            fillColumn();

            $('#helpCallout').hide();
            $('#trim_space').click(trimSpace);
            $('#buttonOk').click(highlightHeader);
            $('#checkbox_all').click(checkCheckbox);
            $('#homeButton').click(redirectHome);

            //Show and hide error message if columns have same header name
            document.getElementById("buttonClose").onclick = function () {
                document.getElementById('showEmbeddedDialog').style.visibility = 'hidden';
            }

            document.getElementById("refresh_icon").onclick = function () {
                window.location = "trim_spaces.html";
            }


            //hide result message
            document.getElementById("resultClose").onclick = function () {
                document.getElementById('resultDialog').style.visibility = 'hidden';
                window.location = "trim_spaces.html";
            }
            document.getElementById("resultOk").onclick = function () {
                document.getElementById('resultDialog').style.visibility = 'hidden';
                window.location = "trim_spaces.html";
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
            worksheet.load('name');
            firstRow.load('address');
            firstCol.load('address');
            worksheet.load('name');

            return ctx.sync().then(function() {

                var tmp_offset = firstCol.address;
                var col_offset = tmp_offset.substring(tmp_offset.indexOf("!") + 1, tmp_offset.indexOf(":"));
                var tmp_row = firstRow.address;
                var row_offset = tmp_row.substring(tmp_row.indexOf("!") + 1, tmp_row.indexOf(":"));

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
            worksheet.load('name');
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
            worksheet.load('name');
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
                            Office.context.document.settings.set('same_header_trim', true);
                        }
                    }
                }

                for (var i = 0; i < range.text[0].length; i++) {
                    if (range.text[0][i] != ""){
                        addNewCheckboxToContainer (range.text[0][i], "column_checkbox" ,"checkboxes_columns");
                    }
                    else {
                        var colchar = getCharFromNumber(i + add_col);
                        addNewCheckboxToContainer ("Column " + colchar, "column_checkbox" ,"checkboxes_columns");
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


    function trimSpace() {

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

                var header = 0;
                var act_worksheet = ctx.workbook.worksheets.getActiveWorksheet();
                var tmp_offset = firstCol.address;
                var col_offset = tmp_offset.substring(tmp_offset.indexOf("!") + 1, tmp_offset.indexOf(":"));
                var tmp_row = firstRow.address;
                var row_offset = Number(tmp_row.substring(tmp_row.indexOf("!") + 1, tmp_row.indexOf(":")));
                var add_col = getNumberFromChar(col_offset);
                var startCell = col_offset + row_offset;

                backupForUndo(range, startCell, add_col, row_offset);

                var checked_checkboxes = getCheckedBoxes("column_checkbox");

                for (var run = 0; run < checked_checkboxes.length; run++) {
                    var trim_array = [];
                    for (var k = 0; k < range.text[0].length; k++) {
                        if (checked_checkboxes[run].id == range.text[0][k] || checked_checkboxes[run].id == "Column " + getCharFromNumber(k + add_col)){
                            header = k;
                            break;
                        }
                    }

                    for (var i = 0; i < range.text.length; i++) {
                        var trim_string = [];
                        trim_string.push(range.text[i][header].trim());
                        trim_array.push(trim_string);
                    }

                    var column_char = getCharFromNumber(header + add_col);
                    var insert_address = column_char + row_offset + ":" + column_char + (range.text.length + row_offset - 1);
                    addTrimArray(trim_array, insert_address);

                }

                if(checked_checkboxes.length == 1) {
                    var endString = " column you seleced."
                } else {
                    var endString = " columns you selected."
                }

                if (document.getElementById('createBackup').checked == true) {
                    var sheet_count = Office.context.document.settings.get('backup_sheet_count') + 1;
                    Office.context.document.settings.set('backup_sheet_count', sheet_count);
                    Office.context.document.settings.saveAsync();
                    var newName = worksheet.name + "(" + sheet_count + ")";
                    addBackupSheet(newName, startCell, add_col, row_offset, function() {
                        var txt = document.createElement("p");
                        txt.className = "ms-font-xs ms-embedded-dialog__content__text";
                        txt.innerHTML = "PrepJet successfully removed all leading and trailing spaces in the " + checked_checkboxes.length + endString;
                        document.getElementById('resultText').appendChild(txt);
                        document.getElementById('resultDialog').style.visibility = 'visible';
                    });

                } else {
                    var txt = document.createElement("p");
                    txt.className = "ms-font-xs ms-embedded-dialog__content__text";
                    txt.innerHTML = "PrepJet successfully removed all leading and trailing spaces in the " + checked_checkboxes.length + endString;
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


    function addTrimArray(trim_array, insert_address){

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange(true);

            range.load('text');
            worksheet.load('name');

            return ctx.sync().then(function() {
                addContentNew(worksheet.name, insert_address, trim_array, function () {});
            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

    }


})();