function displayFieldBegin(){
    if (document.getElementById('beginning_options').value == "custom_b"){
        $('#delimiter_beginning').show();
    }
}
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

            //range.load('address');
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

            if (document.getElementById('ending_options').value == "custom_e"){
                var split_end = document.getElementById('delimiter_input_e').value;
            }
            else {
                var split_end = document.getElementById('ending_options').value;
            }
            if (split_end == "whitespace_e") {
                split_end = " ";
            }

            range.load('text');

            var range_all_adding_to = worksheet.getRange();
            var range_adding_to = range_all_adding_to.getUsedRange();

            range_adding_to.load('address');
            range_adding_to.load('text');

            return ctx.sync().then(function() {
                var header = 0;

                for (var k = 0; k < range.text[0].length; k++){
                    if (selected_identifier == range.text[0][k]){
                        header = k;
                    }
                }

                //todo where to put this??
                var tables = ctx.workbook.tables;
                //var column = tables.getItem("Tabelle1").columns.add(null, extractedValue);
                //column.load('name');

                var act_worksheet = ctx.workbook.worksheets.getActiveWorksheet();
                for (var i = 1; i < range.text.length; i++) {

                    if (split_beginning == "col_beginning"){
                        var position1 = 0;
                    }
                    else {
                        var position1 = range.text[i][header].indexOf(split_beginning);
                    }

                    if (split_end == "col_end") {
                        var position2 = range.text[i][header].length;
                    }
                    else {
                        var position2 = range.text[i][header].indexOf(split_end);
                    }

                    var extractedValue = range.text[i][header].substring(position1, position2);
                    var column_char = getCharFromNumber(1 + range_adding_to.text[0].length);
                    var sheet_row = i + 1;

                    //addContentToWorksheet(act_worksheet, column_char + sheet_row, extractedValue);
                    var column = tables.getItem("Tabelle1").columns.add(null, extractedValue);
                    column.load('name');

                    console.log(column_char + sheet_row)
                }



            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }





    function applyButtonClicked() {
        $('#step1').show();
        $('#step2').hide();
        $('#step3').hide();

        // find columns to match
        var selected_identifier1 = document.getElementById('reference_column_ckeckboxes_1').value; // TODO better reference by ID than name
        var selected_identifier2 = document.getElementById('reference_column_ckeckboxes_1').value; // TODO better reference by ID than name

        var selected_table1 = document.getElementById('table1_options').value; // TODO better reference by ID than name
        var selected_table2 = document.getElementById('table2_options').value; // TODO better reference by ID than name

        Excel.run(function (ctx) {
            var worksheet = ctx.workbook.worksheets.getItem(selected_table2);

            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('address');
            range.load('text');


            var worksheet_adding_to = ctx.workbook.worksheets.getItem(selected_table1);

            var range_all_adding_to = worksheet_adding_to.getRange();
            var range_adding_to = range_all_adding_to.getUsedRange();

            range_adding_to.load('address');
            range_adding_to.load('text');


            return ctx.sync().then(function() {

                // initialize ids
                var sheet1_id = 0;
                var sheet2_id = 0;

                // iterate over columns

                for (var k = 0; k < range.text[0].length; k++){
                    if (selected_identifier1 == range.text[0][k]){
                        sheet1_id = k;
                    }
                }

                for (var k = 0; k < range_adding_to.text[0].length; k++){
                    if (selected_identifier2 == range_adding_to.text[0][k]){
                        sheet2_id = k;
                    }
                }

                for (var k = 0; k < range.text[0].length; k++){

                    // iterate over checked checkboxes

                    var checked_checkboxes = getCheckedBoxes("reference_column_checkbox");

                    for (var l = 0; l < checked_checkboxes.length; l++){ // TODO throws error if none are checked

                        if (checked_checkboxes[l].id == range.text[0][k]){

                            var column_char = getCharFromNumber(1 + l + range_adding_to.text[0].length);

                            // copy title
                            addContentToWorksheet(worksheet_adding_to, column_char + "1", range.text[0][k]);

                            // copy rest
                            for (var i = 1; i < range.text.length; i++) {
                                for (var j = 1; j < range_adding_to.text.length; j++) {
                                    if (range_adding_to.text[j][sheet2_id] == range.text[i][sheet1_id]) {
                                        var sheet_row = j + 1;
                                        addContentToWorksheet(worksheet_adding_to, column_char + sheet_row, range.text[i][k])
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