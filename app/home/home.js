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
            app.initialize();
            fillColumn();

            $('#delimiter_beginning').hide();

            $(".dropdown_table").Dropdown();
            //todo text field not working correctly, placeholder does not disappear on click
            //$(".ms-TextField").TextField();

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


    function splitValue() {
        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();
            var selected_identifier = document.getElementById('column_options').value;

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
                var header = 0;

                for (var k = 0; k < range.text[0].length; k++){
                    if (selected_identifier == range.text[0][k]){
                        header = k;
                    }
                }

                var act_worksheet = ctx.workbook.worksheets.getActiveWorksheet();
                //todo header verschieben
                //var rangeaddress = getCharFromNumber(header + 2) + 1;
                //var range_insert = ctx.workbook.worksheets.getActiveWorksheet().getRange(rangeaddress);
                //range_insert.insert("Right");

                for (var i = 1; i < range.text.length; i++) {

                    //todo loop für alle positions des delimiters
                    var position1 = range.text[i][header].indexOf(delimiter_type);

                    var splitValue1 = range.text[i][header].substring(0, position1);
                    var splitValue2 = range.text[i][header].substring(position1 + 1, range.text[i][header].length);

                    var column_char = getCharFromNumber(header + 2)
                    var sheet_row = i + 1;

                    var rangeaddress = column_char + sheet_row;
                    var range_insert = ctx.workbook.worksheets.getActiveWorksheet().getRange(rangeaddress);
                    range_insert.insert("Right");
                    addContentToWorksheet(act_worksheet, getCharFromNumber(header + 1) + sheet_row, splitValue1);
                    addContentToWorksheet(act_worksheet, getCharFromNumber(header + 2) + sheet_row, splitValue2);

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

})();
