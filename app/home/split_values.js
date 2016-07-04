//display textfield for custom delimiter if selected by user
function displayFieldDelimiter(){
    if (document.getElementById('delimiter_options').value == "custom_delimiter"){
        $('#delimiter_beginning').show();
    }
}

(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            Office.context.document.settings.set('last_clicked_function', "split_values.html");
            if (Office.context.document.settings.get('prepjet_loaded_before') == null) {
                Office.context.document.settings.set('prepjet_loaded_before', true);
                Office.context.document.settings.saveAsync();
                window.location = "intro.html";
            }
            app.initialize();
            fillColumn();

            $('#delimiter_beginning').hide();

            $(".dropdown_table").Dropdown();
            $(".ms-TextField").TextField();

            $('#split_Value').click(splitValue);

        });
    };


    function fillColumn(){

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('text');
            return ctx.sync().then(function() {
                for (var i = 0; i < range.text[0].length; i++) {
                    var el = document.createElement("option");
                    el.value = range.text[0][i];
                    el.textContent = range.text[0][i];
                    document.getElementById("column_options").appendChild(el);
                }
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

            range.load('text');
            var range_all_adding_to = worksheet.getRange();
            var range_adding_to = range_all_adding_to.getUsedRange();
            range_adding_to.load('address');
            range_adding_to.load('text');


            return ctx.sync().then(function() {

                //get column number which to split
                var header = 0;
                for (var k = 0; k < range.text[0].length; k++){
                    if (selected_identifier == range.text[0][k]){
                        header = k;
                    }
                }

                //define variables for array to hold splitted values and length measures
                var act_worksheet = ctx.workbook.worksheets.getActiveWorksheet();
                var array_length = 0;
                var max_array_length = 0;
                var split_array = new Array(range.text.length);


                //loop through whole column, create an array with splitted values and get maximum length
                for (var i = 1; i < range.text.length; i++) {
                    split_array[i] = range.text[i][header].split(delimiter_type);
                    array_length = split_array[i].length
                    if (max_array_length < array_length){
                        max_array_length = array_length
                    }
                }

                //insert empty columns right to split column for splitted parts
                for (var i = 0; i < range.text.length; i++) {
                    for (var j = 1; j < max_array_length; j++) {
                        var column_char = getCharFromNumber(header + 2);
                        var sheet_row = i + 1;
                        var rangeaddress = column_char + sheet_row;
                        var range_insert = ctx.workbook.worksheets.getActiveWorksheet().getRange(rangeaddress);
                        range_insert.insert("Right");
                    }
                }

                //insert splitted parts into new empty columns
                for (var i = 1; i < range.text.length; i++) {
                    var sheet_row = i + 1;
                    for(var j = 0; j < split_array[i].length; j++){
                        addContentToWorksheet(act_worksheet, getCharFromNumber(header + 1 + j) + sheet_row, split_array[i][j]);
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
