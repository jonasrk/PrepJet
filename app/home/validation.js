//show textfield for ending delimiter if custom is selected
function displayBetween(){
    if(document.getElementById('then_operator').value == "between") {
        $('#between_and').show();
    }
}


(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {
            app.initialize();
            fillColumn();

            //$('#delimiter_end').hide();
            $('#between_and').hide();

            $(".dropdown_table").Dropdown();
            $(".ms-TextField").TextField();

            $('#apply').click(validation);

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

                for (var i = 0; i < range.text[0].length; i++) {
                    var el = document.createElement("option");
                    el.value = range.text[0][i];
                    el.textContent = range.text[0][i];
                    document.getElementById("column2_options").appendChild(el);
                }

                $(".dropdown_table_col1").Dropdown();
                $(".dropdown_table_col2").Dropdown();
            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

    }


    function validation() {
        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();
            var selected_identifier1 = document.getElementById('column1_options').value;
            var selected_identifier2 = document.getElementById('column2_options').value;

            //get operator applicable for if condition
            var ifoperator = document.getElementById('if_operator').value;
            if (document.getElementById('if_operator').value == "equal"){
                var ifoperator = "=";
            }
            else if (document.getElementById('if_operator').value == "smaller"){
                var ifoperator = "<";
            }
            else if (document.getElementById('if_operator').value == "greater"){
                var ifoperator = ">";
            }
            else if (document.getElementById('if_operator').value == "inequal"){
                var ifoperator = "!=";
            }
            else { //todo useful return value if nothing is selected
                var ifoperator = 1;
            }


            //get operator applicable for then condition
            var thenoperator = document.getElementById('then_operator').value;
            if (document.getElementById('then_operator').value == "equal"){
                var thenoperator = "=";
            }
            else if (document.getElementById('then_operator').value == "smaller"){
                var thenoperator = "<";
            }
            else if (document.getElementById('then_operator').value == "greater"){
                var thenoperator = ">";
            }
            else if (document.getElementById('then_operator').value == "inequal"){
                var thenoperator = "!=";
            }
            else if (document.getElementById('then_operator').value == "between"){
                var thenoperator = "between";
            }
            else { //todo useful return value if nothing is selected
                var thenoperator = 1;
            }


            //get used range in active Sheet
            range.load('text');
            var range_all_adding_to = worksheet.getRange();
            var range_adding_to = range_all_adding_to.getUsedRange();
            range_adding_to.load('address');
            range_adding_to.load('text');


            return ctx.sync().then(function() {
                var header_if = 0;
                var header_then = 0;

                //get column in header for which to check if condition
                for (var k = 0; k < range.text[0].length; k++){
                    if (selected_identifier1 == range.text[0][k]){
                        header_if = k;
                    }
                }

                //get column in header for which to check then condition
                for (var k = 0; k < range.text[0].length; k++){
                    if (selected_identifier2 == range.text[0][k]){
                        header_then = k;
                    }
                }

                console.log(header_if);
                console.log(header_then);

                //loop through whole column to extract value from
                for (var i = 1; i < range.text.length; i++) {

                    //set position to insert extracted value
                    //var rangeaddress = column_char + sheet_row;
                    //var range_insert = ctx.workbook.worksheets.getActiveWorksheet().getRange(rangeaddress);
                    //range_insert.insert("Right");
                    //addContentToWorksheet(act_worksheet, column_char + sheet_row, extractedValue);

                }


            });

            console.log("Test")
            window.open("extract_values.html","_self");

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }


})();