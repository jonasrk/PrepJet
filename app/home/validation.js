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

            // todo string and number not parsed correclty currently
            var ifcondition = document.getElementById('if_condition').value;
            var thencondition = Number(document.getElementById('then_condition').value);


            //get used range in active Sheet
            range.load('text');
            range.load('valueTypes');
            range.load('values');
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

                //loop through whole column to extract value from
                var act_worksheet = ctx.workbook.worksheets.getActiveWorksheet();
                for (var i = 1; i < range.text.length; i++) {

                    var sheet_row = i + 1;
                    var address = getCharFromNumber(header_then + 1) + sheet_row;

                    if (document.getElementById('if_operator').value == "equal") {
                        if (range.values[i][header_if] == document.getElementById('if_condition').value) {
                            if (document.getElementById('then_operator').value == "equal") {
                                if (range.values[i][header_then] != thencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "smaller") {
                                if (range.values[i][header_then] >= thencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "greater") {
                                if (range.values[i][header_then] <= thencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "inequal") {
                                if (range.values[i][header_then] == thencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "between") {
                                if (range.values[i][header_then] < thencondition || range.values[i][header_then] > document.getElementById('between_and')) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                        }
                    }

                    if (document.getElementById('if_operator').value == "smaller") {
                        var ifcondition = Number(document.getElementById('if_condition').value);
                        if (range.values[i][header_if] < ifcondition) {
                            if (document.getElementById('then_operator').value == "equal") {
                                if (range.values[i][header_then] != thencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "smaller") {
                                if (range.values[i][header_then] >= thencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "greater") {
                                if (range.values[i][header_then] <= thencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "inequal") {
                                if (range.values[i][header_then] == thencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "between") {
                                if (range.values[i][header_then] < thencondition || range.values[i][header_then] > document.getElementById('between_and')) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                        }
                    }

                    if (document.getElementById('if_operator').value == "greater") {
                        var ifcondition = Number(document.getElementById('if_condition').value);
                        if (range.values[i][header_if] > ifcondition) {
                            if (document.getElementById('then_operator').value == "equal") {
                                if (range.values[i][header_then] != thencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "smaller") {
                                if (range.values[i][header_then] >= thencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "greater") {
                                if (range.values[i][header_then] <= thencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "inequal") {
                                if (range.values[i][header_then] == thencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "between") {
                                if (range.values[i][header_then] < thencondition || range.values[i][header_then] > document.getElementById('between_and')) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                        }
                    }
                    if (document.getElementById('if_operator').value == "inequal") {
                        if (range.values[i][header_if] != document.getElementById('if_condition').value) {
                            if (document.getElementById('then_operator').value == "equal") {
                                if (range.values[i][header_then] != thencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "smaller") {
                                if (range.values[i][header_then] >= thencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "greater") {
                                if (range.values[i][header_then] <= thencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "inequal") {
                                if (range.values[i][header_then] == thencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "between") {
                                if (range.values[i][header_then] < thencondition || range.values[i][header_then] > document.getElementById('between_and')) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
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