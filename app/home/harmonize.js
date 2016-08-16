function redirectHome() {
    window.location = "mac_start.html";
}

(function () {
    // 'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            Office.context.document.settings.set('same_header_harmonize', false);
            Office.context.document.settings.set('last_clicked_function', "trim_spaces.html");
            if (Office.context.document.settings.get('prepjet_loaded_before') == null) {
                Office.context.document.settings.set('backup_sheet_count', 1);
                Office.context.document.settings.set('prepjet_loaded_before', true);
                Office.context.document.settings.saveAsync();
                window.location = "intro.html";
            }

            app.initialize();
            fillColumn();


            $(".harmonize_drop").Dropdown();
            $('#harmonize').click(harmonize);
            $('#checkbox_all').click(checkCheckbox);
            $('#buttonOk').click(highlightHeader);
            $('#homeButton').click(redirectHome);

            //Show and hide error message if column have same header name
            document.getElementById("buttonClose").onclick = function () {
                document.getElementById('showEmbeddedDialog').style.visibility = 'hidden';
            }

            document.getElementById("refresh_icon").onclick = function () {
                window.location = "harmonize.html";
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
                                    window.location = "harmonize.html"
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
                    window.location = "harmonize.html";
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
            var range = range_all.getUsedRange();
            var firstCell = range.getColumn(0);
            var firstCol = firstCell.getEntireColumn();
            var tmpRow = range.getRow(0);
            var firstRow = tmpRow.getEntireRow();

            //get used range in active Sheet
            range.load('text');
            firstRow.load('address');
            firstCol.load('address');

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
                            highlightContentInWorksheet(worksheet, getCharFromNumber(run + add_col) + row_offset, '#EA7F04');
                            highlightContentInWorksheet(worksheet, getCharFromNumber(run2 + add_col) + row_offset, '#EA7F04');
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
            var range = range_all.getUsedRange();
            var firstCell = range.getColumn(0);
            var firstCol = firstCell.getEntireColumn();
            var tmpRow = range.getRow(0);
            var firstRow = tmpRow.getEntireRow();

            //get used range in active Sheet
            range.load('text');
            firstRow.load('address');
            firstCol.load('address');

            return ctx.sync().then(function() {

                var tmp_offset = firstCol.address;
                var col_offset = tmp_offset.substring(tmp_offset.indexOf("!") + 1, tmp_offset.indexOf(":"));
                var tmp_row = firstRow.address;
                var row_offset = tmp_row.substring(tmp_row.indexOf("!") + 1, tmp_row.indexOf(":"));
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
            var range = range_all.getUsedRange();
            var firstCell = range.getColumn(0);
            var firstCol = firstCell.getEntireColumn();
            var tmpRow = range.getRow(0);
            var firstRow = tmpRow.getEntireRow();

            //get used range in active Sheet
            range.load('text');
            firstRow.load('address');
            firstCol.load('address');


            return ctx.sync().then(function() {

                var tmp_offset = firstCol.address;
                var col_offset = tmp_offset.substring(tmp_offset.indexOf("!") + 1, tmp_offset.indexOf(":"));
                var tmp_row = firstRow.address;
                var row_offset = tmp_row.substring(tmp_row.indexOf("!") + 1, tmp_row.indexOf(":"));
                var add_col = getNumberFromChar(col_offset);

                for (var run = 0; run < range.text[0].length - 1; run++) {
                    for (var run2 = run + 1; run2 < range.text[0].length; run2++) {
                        if (range.text[0][run] == range.text[0][run2] && range.text[0][run] != "") {
                            document.getElementById('showEmbeddedDialog').style.visibility = 'visible';
                            Office.context.document.settings.set('same_header_harmonize', true);
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


    function harmonize() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();
            var firstCell = range.getColumn(0);
            var firstCol = firstCell.getEntireColumn();
            var tmpRow = range.getRow(0);
            var firstRow = tmpRow.getEntireRow();

            //get used range in active Sheet
            range.load('text');
            worksheet.load('name');
            firstRow.load('address');
            firstCol.load('address');

            var harmo = document.getElementById('harmonize_options').value;

            return ctx.sync().then(function() {

                backupForUndo(range);

                var tmp_offset = firstCol.address;
                var col_offset = tmp_offset.substring(tmp_offset.indexOf("!") + 1, tmp_offset.indexOf(":"));
                var tmp_row = firstRow.address;
                var row_offset = Number(tmp_row.substring(tmp_row.indexOf("!") + 1, tmp_row.indexOf(":")));
                var add_col = getNumberFromChar(col_offset);

                var header = 0;
                var act_worksheet = ctx.workbook.worksheets.getActiveWorksheet();
                var checked_checkboxes = getCheckedBoxes("column_checkbox");

                for (var run = 0; run < checked_checkboxes.length; run++) {
                    for (var k = 0; k < range.text[0].length; k++) {
                        if (checked_checkboxes[run].id == range.text[0][k] || checked_checkboxes[run].id == "Column " + getCharFromNumber(k + add_col)){
                            header = k;
                            break;
                        }
                    }

                    var harm_array = [];

                    for (var k = 0; k < range.text.length; k++) {
                        if (harmo == "allupper") {
                            var harm_string = [range.text[k][header].toUpperCase()];
                        }
                        if (harmo == "alllower") {
                            var harm_string = [range.text[k][header].toLowerCase()];
                        }
                        if (harmo == "firstupper") { //todo when leading space first real letter not transformed
                            var tmp = range.text[k][header].toLowerCase().split(" ");
                            var tmp_upper = [];
                            for (var runtmp = 0; runtmp < tmp.length; runtmp++) {
                                tmp_upper.push(tmp[runtmp].charAt(0).toUpperCase() + tmp[runtmp].slice(1));
                            }
                            var harm_string = [tmp_upper[0]];
                            for (var runtmp = 1; runtmp < tmp_upper.length; runtmp++) {
                                harm_string = [harm_string.concat(" ", tmp_upper[runtmp])];
                            }
                        }
                        if (harmo == "oneupper") {
                            var tmp = range.text[k][header].split(" ");
                            var tmp_upper = [];
                            tmp_upper.push(tmp[0].charAt(0).toUpperCase() + tmp[0].slice(1).toLowerCase());
                            for (var runtmp = 1; runtmp < tmp.length; runtmp++) {
                                tmp_upper.push(tmp[runtmp].charAt(0) + tmp[runtmp].slice(1).toLowerCase());
                            }

                            var harm_string = [tmp_upper[0]];
                            for (var runtmp = 1; runtmp < tmp_upper.length; runtmp++) {
                                harm_string = [harm_string.concat(" ", tmp_upper[runtmp])];
                            }
                        }

                        harm_array.push(harm_string);

                    }

                    var insert_address = getCharFromNumber(header + add_col) + row_offset + ":" + getCharFromNumber(header + add_col) + (range.text.length + row_offset - 1);
                    addContentNew(worksheet.name, insert_address, harm_array, function () {});

                    var i = 0;

                    if (document.getElementById('createBackup').checked != true) {
                        addContentNew(worksheet.name, insert_address, harm_array, function () {
                            i++;
                            if (i >= checked_checkboxes.length){
                                window.location = "harmonize.html";
                            }
                        });
                    } else {
                        addContentNew(worksheet.name, insert_address, harm_array, function () {});
                    }
                }

                var startCell = col_offset + row_offset;
                if (document.getElementById('createBackup').checked == true) {
                    var sheet_count = Office.context.document.settings.get('backup_sheet_count') + 1;
                    Office.context.document.settings.set('backup_sheet_count', sheet_count);
                    Office.context.document.settings.saveAsync();
                    var newName = worksheet.name + "(" + sheet_count + ")";
                    addBackupSheet(newName, startCell, add_col, function () {
                        window.location = "harmonize.html"
                    });

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