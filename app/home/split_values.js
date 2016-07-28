function displayAdvancedCount() {
        $('#delimiter_count').show();
        $('.delimiter_count_dropdown').show();
        $('#advanced_settings').hide();
        $('#advanced_hide').show();
        Office.context.document.settings.set('more_option', true);
}

function hideAdvancedCount() {
        $('#delimiter_count').hide();
        $('.delimiter_count_dropdown').hide();
        $('#advanced_settings').show();
        $('#advanced_hide').hide();
        Office.context.document.settings.set('more_option', false);
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

(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            Office.context.document.settings.set('same_header_split', false);
            Office.context.document.settings.set('more_option', false);
            Office.context.document.settings.set('last_clicked_function', "split_values.html");
            if (Office.context.document.settings.get('prepjet_loaded_before') == null) {
                Office.context.document.settings.set('prepjet_loaded_before', true);
                Office.context.document.settings.saveAsync();
                window.location = "intro.html";
            }
            app.initialize();
            fillColumn();

            if (Office.context.document.settings.get('same_header_split') == false) {
                $("#showEmbeddedDialog").hide();
            }

            $('#delimiter_beginning').hide();
            $('#delimiter_count').hide();
            $('#checkbox_delimiter').hide();
            $(".delimiter_count_dropdown").Dropdown().hide()
            $(".keep_delimiter_dropdown").Dropdown().hide()
            $('#advanced_hide').hide();

            $(".dropdown_table").Dropdown();
            $(".ms-TextField").TextField();

            $('#split_Value').click(splitValue);
            $('#advanced_settings').click(displayAdvancedCount);
            $('#advanced_hide').click(hideAdvancedCount);


            // Hides the dialog.
            document.getElementById("buttonClose").onclick = function () {
                $("#showEmbeddedDialog").hide();
            }

            // Performs the action and closes the dialog.
            document.getElementById("buttonOk").onclick = function () {
                $("#showEmbeddedDialog").hide();
            }

            Excel.run(function (ctx) {

                var myBindings = Office.context.document.bindings;
                var worksheetname = ctx.workbook.worksheets.getActiveWorksheet();

                worksheetname.load('name')

                return ctx.sync().then(function() {

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
            });

        });
    };


    function fillColumn(){

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('text');

            return ctx.sync().then(function() {

                for (var run = 0; run < range.text[0].length - 1; run++) {
                    for (var run2 = run + 1; run2 < range.text[0].length; run2++) {
                        if (range.text[0][run] == range.text[0][run2]) {
                            $("#showEmbeddedDialog").show();
                            highlightContentInWorksheet(worksheet, getCharFromNumber(run) + 1, '#EA7F04');
                            highlightContentInWorksheet(worksheet, getCharFromNumber(run2) + 1, '#EA7F04');
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
                        el.value = "Column " + getCharFromNumber(i);
                        el.textContent = "Column " + getCharFromNumber(i);
                    }
                    document.getElementById("column_options").appendChild(el);
                }
                // console.log($(".dropdown_table_col"));
                $(".dropdown_table_col").Dropdown();
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
            var range = range_all.getUsedRange();
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
            if (delimiter_type == "semikolon") {
                delimiter_type = ";";
            }

            //if advanced settings are selected, get values for delimiter count
            if (Office.context.document.settings.get('more_option') == true) {
                var count_delimiter = Number(document.getElementById('delimiter_count_i').value);
                var count_direction = document.getElementById('delimiter_count_drop').value;
            }

            range.load('text');
            var range_all_adding_to = worksheet.getRange();
            var range_adding_to = range_all_adding_to.getUsedRange();
            range_adding_to.load('address');
            range_adding_to.load('text');


            return ctx.sync().then(function() {

                backupForUndo(range_adding_to);

                //get column number which to split
                var header = 0;
                for (var k = 0; k < range.text[0].length; k++){
                    if (selected_identifier == range.text[0][k] || selected_identifier == "Column " + getCharFromNumber(k)){
                        header = k;
                    }
                }

                //define variables for array to hold splitted values and length measures
                var act_worksheet = ctx.workbook.worksheets.getActiveWorksheet();
                var array_length = 0;
                var max_array_length = 0;
                var split_array = new Array(range.text.length);
                var split_array_test = new Array(range.text.length);


                //loop through whole column, create an array with splitted values and get maximum length
                if (Office.context.document.settings.get('more_option') == false) {
                    for (var i = 0; i < range.text.length; i++) {
                        if (range.text[i][header] != "") {
                            split_array[i] = range.text[i][header].split(delimiter_type);
                            array_length = split_array[i].length;
                            if (max_array_length < array_length){
                                max_array_length = array_length;
                            }
                        }
                    }
                }
                else {
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
                                count_delimiter = Number(document.getElementById('delimiter_count_i').value);
                                max_array_length = 2;
                            }
                        }
                        else if (range.text[i][header].indexOf(delimiter_type) == -1) {
                            split_array[i] = range.text[i][header];
                        }
                    }
                }

                //insert empty columns right to split column for splitted parts
                for (var i = 0; i < range.text.length; i++) {
                    for (var j = 1; j < max_array_length; j++) {
                        var column_char = getCharFromNumber(header + 1);
                        var sheet_row = i + 1;
                        var rangeaddress = column_char + sheet_row;
                        var range_insert = ctx.workbook.worksheets.getActiveWorksheet().getRange(rangeaddress);
                        range_insert.insert("Right");
                    }
                }

                //insert splitted parts into new empty columns
                for (var i = 0; i < range.text.length; i++) {
                    var sheet_row = i + 1;
                    if (range.text[i][header] != "" && range.text[i][header].indexOf(delimiter_type) != -1) {
                        for(var j = 0; j < split_array[i].length; j++){
                            addContentToWorksheet(act_worksheet, getCharFromNumber(header + j) + sheet_row, split_array[i][j]);
                        }
                    }
                }
                window.location = "split_values.html";
            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

})();
