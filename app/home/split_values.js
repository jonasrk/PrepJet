//show textfield for beginning delimiter if custom is selected
function displayFieldBegin(){
    if (document.getElementById('beginning_options').value == "custom_b"){
        $('#delimiter_beginning').show();
    }
}

//show textfield for ending delimiter if custom is selected
function displayFieldEnd(){
    if(document.getElementById('ending_options').value == "custom_e") {
        $('#delimiter_end').show();
    }
}


(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {
            app.initialize();
            fillColumn();

            $('#delimiter_end').hide();
            $('#delimiter_beginning').hide();

            $(".dropdown_table").Dropdown();
            $(".ms-TextField").TextField();

            $('#extract_Value').click(extractValue);

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
                    document.getElementById("column1_options").appendChild(el);
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


    function extractValue() {
        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();
            var selected_identifier = document.getElementById('column1_options').value;

            //get character where to start extracting and translate string into delimiter
            var split_beginning = document.getElementById('beginning_options').value;
            if (document.getElementById('beginning_options').value == "custom_b"){
                var split_beginning = document.getElementById('delimiter_input_b').value;
            }
            else {
                var split_beginning = document.getElementById('beginning_options').value;
            }
            if (split_beginning == "whitespace_b") {
                split_beginning = " ";
            }

            //get character where to end extracting and translate string into delimiter
            if (document.getElementById('ending_options').value == "custom_e"){
                var split_end = document.getElementById('delimiter_input_e').value;
            }
            else {
                var split_end = document.getElementById('ending_options').value;
            }
            if (split_end == "whitespace_e") {
                split_end = " ";
            }

            //get (optional) column where to insert extracted value, default is to the right of original column
            var target_column = document.getElementById('target_column_input').value


            //get used range in active Sheet
            range.load('text');
            var range_all_adding_to = worksheet.getRange();
            var range_adding_to = range_all_adding_to.getUsedRange();
            range_adding_to.load('address');
            range_adding_to.load('text');


            return ctx.sync().then(function() {
                var header = 0;

                //get column in header from which to extract value
                for (var k = 0; k < range.text[0].length; k++){
                    if (selected_identifier == range.text[0][k]){
                        header = k;
                    }
                }

                //insert empty cell into header column
                var act_worksheet = ctx.workbook.worksheets.getActiveWorksheet();
                if (target_column != ""){
                    var custom_range_address = target_column + 1;
                    var range_insert = ctx.workbook.worksheets.getActiveWorksheet().getRange(custom_range_address);
                    range_insert.insert("Right");
                }
                else {
                    var rangeaddress = getCharFromNumber(header + 2) + 1;
                    var range_insert = ctx.workbook.worksheets.getActiveWorksheet().getRange(rangeaddress);
                    range_insert.insert("Right");
                }


                //loop through whole column to extract value from
                for (var i = 1; i < range.text.length; i++) {

                    //get index where to start extracting value
                    if (split_beginning == "col_beginning"){
                        var position1 = 0;
                    }
                    else {
                        var position1 = range.text[i][header].indexOf(split_beginning);
                    }

                    //get index where to end extracting value
                    if (split_end == "col_end") {
                        var position2 = range.text[i][header].length;
                    }
                    else {
                        var position2 = range.text[i][header].indexOf(split_end);
                    }

                    //get position where to insert extracted value
                    var sheet_row = i + 1;
                    if (target_column != "") {
                        var column_char = target_column
                    }
                    else {
                        var column_char = getCharFromNumber(header + 2);
                    }

                    //get value to extract
                    var extractedValue = range.text[i][header].substring(position1, position2);

                    //set position to insert extracted value
                    var rangeaddress = column_char + sheet_row;
                    var range_insert = ctx.workbook.worksheets.getActiveWorksheet().getRange(rangeaddress);
                    range_insert.insert("Right");
                    addContentToWorksheet(act_worksheet, column_char + sheet_row, extractedValue);

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