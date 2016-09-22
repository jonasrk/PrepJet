function displayAdvancedCount() {
        $('#delimiter_count').show();
        $('.delimiter_count_dropdown').show();
        $('#split_Value').show();
        //$('#advanced_settings').hide();
        //$('#advanced_hide').show();
}

function hideAdvancedCount() {
        $('#delimiter_count').hide();
        $('.delimiter_count_dropdown').hide();
        $('#advanced_settings').show();
        $('#advanced_hide').hide();
}


//display textfield for custom delimiter if selected by user
function displayFieldDelimiter(){
    if (document.getElementById('delimiter_options').value == "custom_delimiter"){
        $('#delimiter_beginning').show();
    }
    else {
        $('#delimiter_beginning').hide();
    }
}


function redirectHome() {
    window.location = "mac_start.html";
}

function redirectExtract() {
    window.location = "extract_values.html";
}

function backToOne() {
    $('#content-header2').hide();
    $('#step1').show();
    $('#step2').hide();
    $('#step3').hide();
    $('#step4').hide();
}

function backToTwo() {
    $('#content-header2').show();
    $('#step1').hide();
    $('#step2').show();
    $('#step3').hide();
    $('#step4').hide();
}

function backToThree() {
    $('#content-header2').show();
    $('#step1').hide();
    $('#step2').hide();
    $('#step3').show();
    $('#step4').hide();
}

(function () {
    // 'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            Office.context.document.settings.set('same_header_split', false);
            Office.context.document.settings.set('last_clicked_function', "split_values.html");
            if (Office.context.document.settings.get('prepjet_loaded_before') == null) {
                Office.context.document.settings.set('backup_sheet_count', 1);
                Office.context.document.settings.set('prepjet_loaded_before', true);
                Office.context.document.settings.saveAsync();
                window.location = "intro.html";
            }

            app.initialize();
            fillColumn();

            $('#content-header2').hide();
            $('#step2').hide();
            $('#step3').hide();
            $('#step4').hide();

            $('#delimiter_beginning').hide();
            $('#split_Value').hide();
            $('#delimiter_count').Dropdown().hide();
            $('#checkbox_delimiter').hide();
            $(".delimiter_count_dropdown").Dropdown().hide();

            $(".dropdown_table").Dropdown();
            $(".ms-TextField").TextField();

            $('#continue1').click(step3Show);
            $('#continue2').click(step4Show);
            $('#back1').click(backToOne);
            $('#back2').click(backToTwo);
            $('#back3').click(backToThree);

            $('#extractButton').click(redirectExtract);
            $('#splitButton').click(step2Show);
            $('#split_Value').click(splitValue);
            $('#splitApply1').click(splitValue);
            $('#splitApply2').click(displayAdvancedCount);
            $('#buttonOk').click(highlightHeader);
            $('#homeButton').click(redirectHome);
            $('#homeButton2').click(redirectHome);


            // Hides the dialog.
            document.getElementById("buttonClose").onclick = function () {
                document.getElementById('showEmbeddedDialog').style.visibility = 'hidden';
            }


            //Show and hide error message if columns have same header name
            /*document.getElementById("help_icon").onclick = function () {
                document.getElementById('helpCallout').style.visibility = 'visible';
            }
            document.getElementById("closeCallout").onclick = function () {
                document.getElementById('helpCallout').style.visibility = 'hidden';
            }*/

            document.getElementById("refresh_icon").onclick = function () {
                window.location = "split_values.html";
            }

            document.getElementById("resultClose").onclick = function () {
                document.getElementById('resultDialog').style.visibility = 'hidden';
                window.location = "split_values.html";
            }
            document.getElementById("resultOk").onclick = function () {
                document.getElementById('resultDialog').style.visibility = 'hidden';
                window.location = "split_values.html";
            }


            /*Excel.run(function (ctx) {

                //var myBindings = Office.context.document.bindings;
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
                                    window.location = "split_values.html"
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
                    window.location = "split_values.html";
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

    function step2Show() {

        $('#step1').hide();
        $('#content-header2').show();
        $('#step2').show();
        $('#step3').hide();
        $('#step4').hide();

        Excel.run(function (ctx) {

            return ctx.sync().then(function() {

            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function step3Show() {

        $('#content-header2').show();
        $('#step1').hide();
        $('#step2').hide();
        $('#step3').show();
        $('#step4').hide();

        Excel.run(function (ctx) {

            return ctx.sync().then(function() {

            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }


    function step4Show() {

        $('#content-header2').show();
        $('#step1').hide();
        $('#step2').hide();
        $('#step3').hide();
        $('#step4').show();

        Excel.run(function (ctx) {

            return ctx.sync().then(function() {

            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }


    //function to split values in a column by a specified delimiter into different columns
    function splitValue() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange(true);
            var firstCell = range.getColumn(0);
            var firstCol = firstCell.getEntireColumn();
            var tmpRow = range.getRow(0);
            var firstRow = tmpRow.getEntireRow();

            var selected_identifier = document.getElementById('column_options').value;

            //get delimiter where to split and translate user input into delimiter character
            var delimiter_type = document.getElementById('delimiter_options').value;
            if (delimiter_type == "custom_delimiter"){
                var delimiter_type = document.getElementById('delimiter_input').value;
            }
            if (delimiter_type == "whitespace") {
                delimiter_type = " ";
            }
            if (delimiter_type == "comma") {
                delimiter_type = ",";
            }
            if (delimiter_type == "semicolon") {
                delimiter_type = ";";
            }

            range.load('text');
            worksheet.load('name');
            firstRow.load('address');
            firstCol.load('address');

            var range_all_adding_to = worksheet.getRange();
            var range_adding_to = range_all_adding_to.getUsedRange(true);
            range_adding_to.load('address');
            range_adding_to.load('text');

            return ctx.sync().then(function() {

                var tmp_offset = firstCol.address;
                var col_offset = tmp_offset.substring(tmp_offset.indexOf("!") + 1, tmp_offset.indexOf(":"));
                var tmp_row = firstRow.address;
                var row_offset = Number(tmp_row.substring(tmp_row.indexOf("!") + 1, tmp_row.indexOf(":")));
                var add_col = getNumberFromChar(col_offset);
                var startCell = col_offset + row_offset;

                backupForUndo(range, startCell, add_col, row_offset);

                function getCountDelimiter () {
                    var count_delimiter = 0;
                    if (document.getElementById('delimiter_count_i').value == "one") { count_delimiter = 1; }
                    else if(document.getElementById('delimiter_count_i').value == "two") { count_delimiter = 2; }
                    else if(document.getElementById('delimiter_count_i').value == "three") { count_delimiter = 3; }
                    else if(document.getElementById('delimiter_count_i').value == "four") { count_delimiter = 4; }
                    else if(document.getElementById('delimiter_count_i').value == "five") { count_delimiter = 5; }
                    else if(document.getElementById('delimiter_count_i').value == "six") { count_delimiter = 6; }
                    else if(document.getElementById('delimiter_count_i').value == "seven") { count_delimiter = 7; }
                    else if(document.getElementById('delimiter_count_i').value == "eight") { count_delimiter = 8; }
                    else if(document.getElementById('delimiter_count_i').value == "nine") { count_delimiter = 9; }
                    else if(document.getElementById('delimiter_count_i').value == "all") {
                        count_delimiter = 0;
                    }
                    return count_delimiter;
                }

                //get column number which to split
                var header = 0;
                for (var k = 0; k < range.text[0].length; k++){
                    if (selected_identifier == range.text[0][k] || selected_identifier == "Column " + getCharFromNumber(k + add_col)){
                        header = k;
                    }
                }

                //define variables for array to hold splitted values and length measures
                var array_length = 0;
                var max_array_length = 0;
                var split_array = new Array(range.text.length);

                //loop through whole column, create an array with splitted values and get maximum length
                var count_delimiter = getCountDelimiter();
                if (count_delimiter == 0) {
                    for (var i = 0; i < range.text.length; i++) {
                        if (range.text[i][header] != "") {
                            split_array[i] = range.text[i][header].split(delimiter_type);
                            array_length = split_array[i].length;
                            if (max_array_length < array_length){
                                max_array_length = array_length;
                            }
                        }
                        else {
                            split_array[i] = [""];
                        }
                    }

                    for (var j = 0; j < range.text.length; j++) {
                        if (split_array[j].length < max_array_length) {
                            for (var i = split_array[j].length; i < max_array_length; i++) {
                                split_array[j][i] = "";
                            }
                        }
                    }
                }

                else {
                    var count_direction = document.getElementById('delimiter_count_drop').value;
                    for (var i = 0; i < range.text.length; i++) {
                        if (range.text[i][header] != "" && range.text[i][header].indexOf(delimiter_type) != -1) {
                            split_array[i] = range.text[i][header].split(delimiter_type);
                            array_length = split_array[i].length;

                            if (max_array_length < array_length){
                                max_array_length = array_length;
                            }

                            if (count_delimiter != 0) {
                                if (count_direction == "right") {
                                    count_delimiter = array_length - count_delimiter;
                                    if (count_delimiter < 1) {
                                        count_delimiter = array_length;
                                    }
                                }
                                if (count_direction == "left" && (array_length <= count_delimiter)) {
                                    count_delimiter = array_length;
                                }

                                var str1_tmp = split_array[i][0];

                                for (var j = 1; j < count_delimiter; j++) {
                                    str1_tmp = str1_tmp.concat(delimiter_type, split_array[i][j]);
                                }

                                var str2_tmp = split_array[i][count_delimiter];
                                for (var j = count_delimiter + 1; j < array_length; j++) {
                                    str2_tmp = str2_tmp.concat(delimiter_type, split_array[i][j]);
                                }

                                split_array[i] = [str1_tmp];
                                split_array[i].push(str2_tmp);
                                max_array_length = 2;
                            }
                        }
                        else if (range.text[i][header].indexOf(delimiter_type) == -1) {
                            split_array[i] = [range.text[i][header]];
                            split_array[i].push("");
                        }
                    }
                }

                //insert empty columns right to split column for splitted parts
                for (var j = 1; j < max_array_length; j++) {
                    var column_char = getCharFromNumber(header + add_col + 1);
                    var rangeaddress = column_char + ":" + column_char;
                    var range_insert = ctx.workbook.worksheets.getActiveWorksheet().getRange(rangeaddress);
                    range_insert.insert("Right");
                }

                var insert_address = getCharFromNumber(header + add_col) + row_offset + ":" + getCharFromNumber(header + add_col + max_array_length - 1) + (range.text.length + row_offset - 1);
                addSplitValue(split_array, insert_address);

                if (document.getElementById('createBackup').checked == true) {
                    var sheet_count = Office.context.document.settings.get('backup_sheet_count') + 1;
                    Office.context.document.settings.set('backup_sheet_count', sheet_count);
                    Office.context.document.settings.saveAsync();
                    var newName = worksheet.name + "(" + sheet_count + ")";
                    addBackupSheet(newName, startCell, add_col, row_offset, function() {
                        var txt = document.createElement("p");
                        txt.className = "ms-font-xs ms-embedded-dialog__content__text";
                        txt.innerHTML = "PrepJet successfully splitted your data."
                        document.getElementById('resultText').appendChild(txt);

                        document.getElementById('resultDialog').style.visibility = 'visible';
                    });
                }
                else {
                    var txt = document.createElement("p");
                    txt.className = "ms-font-xs ms-embedded-dialog__content__text";
                    txt.innerHTML = "PrepJet successfully splitted your data."
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


    function addSplitValue(split_array, insert_address){

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange(true);

            range.load('text');
            worksheet.load('name');

            return ctx.sync().then(function() {
                    addContentNew(worksheet.name, insert_address, split_array, function () {});
            });
        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

    }

})();
