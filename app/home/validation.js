//display fields for advanced rule
function displayAdvanced(){
    $('#simple_button').show();
    $('#simple_dropdown').hide();
    $('#advanced_button').hide();
    $('#advanced_dropdown').show();
    $('#advanced_dropdown2').show();
    $('#then_op_drop').show();
    $('#delimiter_end').show();
    $('#apply_advanced').show();
    $('#apply_simple').hide();
}

//display fields for simple rule
function displaySimple() {
            $('#advanced_button').hide();
            $('#simple_button').hide();
            $('#simple_dropdown').show();
            $('#advanced_dropdown').hide();
            $('#advanced_dropdown2').hide();
            $('#then_op_drop').hide();
            $('#delimiter_end').hide();
            $('#apply_advanced').hide();
            $('#apply_advanced').show();
}

//show textfield for ending delimiter if custom is selected
function displayBetween(){
    if(document.getElementById('then_operator').value == "between" || document.getElementById('then_operator').value == "notbetween") {
        $('#betweenand').show();
    }
    else {
        $('#betweenand').hide();
    }
}

var activeSelection = 0;
function setFocus(activeID) {
    activeSelection = activeID;
}

function displaySimpleBetween(){
    if(document.getElementById('if_operator').value == "between" || document.getElementById('if_operator').value == "notbetween") {
        $('#between_beginning').show();
    }
    else {
        $('#between_beginning').hide();
    }
}


(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            Office.context.document.settings.set('last_clicked_function', "validation.html");
            if (Office.context.document.settings.get('prepjet_loaded_before') == null) {
                Office.context.document.settings.set('prepjet_loaded_before', true);
                Office.context.document.settings.saveAsync();
                window.location = "intro.html";
            }

            app.initialize();
            fillSimpleColumn();

            $('#simple_button').hide();
            $('#advanced_dropdown').hide();
            $('#advanced_dropdown2').hide();
            $('#then_op_drop').hide();
            $('#delimiter_end').hide();
            $('#apply_advanced').hide();

            fillColumn();
            $('#advanced_button').click(displayAdvanced);
            $('#simple_button').click(displaySimple);

            $('#betweenand').hide();
            $('#between_beginning').hide();
            $(".dropdown_table").Dropdown();
            $(".ms-TextField").TextField();

            $('#apply_advanced').click(validationAdvanced);
            $('#apply_simple').click(validationSimple);


            Office.context.document.addHandlerAsync("documentSelectionChanged", myIfHandler, function(result){}
            );
            // Event handler function.
            function myIfHandler(eventArgs){
                Excel.run(function (ctx) {
                    var selectedRange = ctx.workbook.getSelectedRange();
                    selectedRange.load('text');
                    return ctx.sync().then(function () {
                        if (activeSelection == 0) {
                            writeif(selectedRange.text);
                        }
                        else if (activeSelection == 1) {
                            writeifand(selectedRange.text);
                        }
                        else if (activeSelection == 2) {
                            writethen(selectedRange.text);
                        }
                        else if (activeSelection == 3) {
                            writethenand(selectedRange.text);
                        }
                    });
                });
            }
            // Function that writes to a div with id='message' on the page.
            function writeif(message){
                document.getElementById('if_condition').value = message;
            }

            function writeifand(message){
                document.getElementById('if_between_condition').value = message;
            }

            function writethen(message){
                document.getElementById('then_condition').value = message;
            }

            function writethenand(message){
                document.getElementById('between_and').value = message;
            }

        });
    };

    function fillSimpleColumn(){

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();
            range.load('text');
            return ctx.sync().then(function() {
                for (var i = 0; i < range.text[0].length; i++) {
                    var el = document.createElement("option");
                    if (range.text[0][i] != "") {
                            el.value = range.text[0][i];
                            el.textContent = range.text[0][i];
                        }
                        else {
                            el.value = "Column " + getCharFromNumber(i + 1);
                            el.textContent = "Column " + getCharFromNumber(i + 1);
                        }
                    document.getElementById("column_simple").appendChild(el);
                }

                $(".table_simple").Dropdown();
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
            range.load('text');
            return ctx.sync().then(function() {
                for (var i = 0; i < range.text[0].length; i++) {
                    var el = document.createElement("option");
                    if (range.text[0][i] != "") {
                            el.value = range.text[0][i];
                            el.textContent = range.text[0][i];
                        }
                        else {
                            el.value = "Column " + getCharFromNumber(i + 1);
                            el.textContent = "Column " + getCharFromNumber(i + 1);
                        }
                    document.getElementById("column1_options").appendChild(el);
                }

                for (var i = 0; i < range.text[0].length; i++) {
                    var el = document.createElement("option");
                    if (range.text[0][i] != "") {
                            el.value = range.text[0][i];
                            el.textContent = range.text[0][i];
                        }
                        else {
                            el.value = "Column " + getCharFromNumber(i + 1);
                            el.textContent = "Column " + getCharFromNumber(i + 1);
                        }
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


    //validation when simple rule is created
    function validationSimple() {
        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();
            var selected_identifier = document.getElementById('column_simple').value;


            if (document.getElementById('if_operator').value == "inlist") {
                var in_list = document.getElementById('if_condition').value;
                var splitted_list = in_list.split(",");
                for (var run = 0; run < splitted_list.length; run ++) {
                    splitted_list[run] = splitted_list[run].trim();
                }
                for (var run = 0; run < splitted_list.length; run++) {
                    if (isNaN(Number(splitted_list[run])) != true) {
                        splitted_list[run] = Number(splitted_list[run]);
                    }
                }
            }
            else {
                if (isNaN(Number(document.getElementById('if_condition').value)) == true) {
                    var ifcondition = document.getElementById('if_condition').value;
                }
                else {
                    var ifcondition = Number(document.getElementById('if_condition').value);
                }
            }

            if (document.getElementById('if_operator').value == "notbetween" || document.getElementById('if_operator').value == "between") {
                if (isNaN(Number(document.getElementById('if_between_condition').value)) == true) {
                    var ifbetweencondition = document.getElementById('if_between_condition').value;
                }
                else {
                    var ifbetweencondition = Number(document.getElementById('if_between_condition').value);
                }
            }


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

                //get column in header for which to check if condition
                for (var k = 0; k < range.text[0].length; k++){
                    if (selected_identifier == range.text[0][k] || selected_identifier == "Column " + getCharFromNumber(k + 1)){
                        header_if = k;
                    }
                }

                //loop through whole column to extract value from
                var act_worksheet = ctx.workbook.worksheets.getActiveWorksheet();
                for (var i = 1; i < range.text.length; i++) {

                    var sheet_row = i + 1;
                    var address = getCharFromNumber(header_if + 1) + sheet_row;

                    if (document.getElementById('if_operator').value == "equal") {
                        if (range.values[i][header_if] != ifcondition) {
                            highlightContentInWorksheet(act_worksheet, address, "red");
                        }
                    }

                    if (document.getElementById('if_operator').value == "smaller") {
                        if (range.values[i][header_if] >= ifcondition) {
                            highlightContentInWorksheet(act_worksheet, address, "red");
                        }
                    }

                    if (document.getElementById('if_operator').value == "greater") {
                        if (range.values[i][header_if] <= ifcondition) {
                            highlightContentInWorksheet(act_worksheet, address, "red");
                        }
                    }

                    if (document.getElementById('if_operator').value == "inequal") {
                        if (range.values[i][header_if] == ifcondition) {
                             highlightContentInWorksheet(act_worksheet, address, "red");
                        }
                    }

                    if (document.getElementById('if_operator').value == "between") {
                        if (range.values[i][header_if] < ifcondition || range.values[i][header_if] > ifbetweencondition) {
                             highlightContentInWorksheet(act_worksheet, address, "red");
                        }
                    }

                    if (document.getElementById('if_operator').value == "notbetween") {
                        if (range.values[i][header_if] > ifcondition && range.values[i][header_if] < ifbetweencondition) {
                             highlightContentInWorksheet(act_worksheet, address, "red");
                        }
                    }

                    if (document.getElementById('if_operator').value == "inlist") {
                        var check = 0;
                        for (run = 0; run < splitted_list.length; run++) {
                            if (range.values[i][header_if] == splitted_list[run]) {
                                 check = 1;
                            }
                        }
                        if (check != 1) {
                            highlightContentInWorksheet(act_worksheet, address, "red");
                        }
                    }
                }
                window.open("validation.html");
            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }


    //validation when advanced rule is selected
    function validationAdvanced() {
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
            else if (document.getElementById('then_operator').value == "notbetween") {
                var thenoperator = "notbetween";
            }
            else if (document.getElementById('then_operator').value == "inlist") {
                var in_then_list = document.getElementById('then_condition').value;
                var splitted_then_list = in_then_list.split(",");
                for (var run = 0; run < splitted_then_list.length; run ++) {
                    splitted_then_list[run] = splitted_then_list[run].trim();
                }
                for (var run = 0; run < splitted_then_list.length; run ++) {
                    if (isNaN(Number(splitted_then_list[run])) != true) {
                        splitted_then_list[run] = Number(splitted_then_list[run]);
                    }
                }
            }
            else { //todo useful return value if nothing is selected
                var thenoperator = 1;
            }

            //get correct list with values entered by user
            if (document.getElementById('if_operator').value == "inlist") {
                var in_if_list = document.getElementById('if_condition').value;
                var splitted_if_list = in_if_list.split(",");
                for (var run = 0; run < splitted_if_list.length; run ++) {
                    splitted_if_list[run] = splitted_if_list[run].trim();
                }
                for (var run = 0; run < splitted_if_list.length; run++) {
                    if (isNaN(Number(splitted_if_list[run])) != true) {
                        splitted_if_list[run] = Number(splitted_if_list[run]);
                    }
                }
            }
            //get correct value for condition in if statement
            else {
                if (isNaN(Number(document.getElementById('if_condition').value)) == true) {
                    var ifcondition = document.getElementById('if_condition').value;
                }
                else {
                    var ifcondition = Number(document.getElementById('if_condition').value);
                }
            }

            //get correct value in if condition for between/not between 2nd value
            if (document.getElementById('if_operator').value == "between" || document.getElementById('if_operator').value == "notbetween") {
                if (isNaN(Number(document.getElementById('if_between_condition').value)) == true) {
                    var ifbetweencondition = document.getElementById('if_between_condition').value;
                }
                else {
                    var ifbetweencondition = Number(document.getElementById('if_between_condition').value);
                }
            }

            //get correct value in then condition
            if (document.getElementById('then_operator').value != "inlist") {
                if (isNaN(Number(document.getElementById('then_condition').value)) == true) {
                    var thencondition = document.getElementById('then_condition').value;
                }
                else {
                    var thencondition = Number(document.getElementById('then_condition').value);
                }
            }

            //get correct value in then condition for between/not between 2nd value
            if (document.getElementById('then_operator').value == "between" || document.getElementById('then_operator').value == "notbetween") {
                if (isNaN(Number(document.getElementById('between_and').value)) == true) {
                    var betweencondition = document.getElementById('between_and').value;
                }
                else {
                    var betweencondition = Number(document.getElementById('between_and').value);
                }
            }


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
                    if (selected_identifier1 == range.text[0][k] || selected_identifier1 == "Column " + getCharFromNumber(k + 1)){
                        header_if = k;
                    }
                }

                //get column in header for which to check then condition
                for (var k = 0; k < range.text[0].length; k++){
                    if (selected_identifier2 == range.text[0][k] || selected_identifier2 == "Column " + getCharFromNumber(k + 1)){
                        header_then = k;
                    }
                }

                //loop through whole column to extract value from
                var act_worksheet = ctx.workbook.worksheets.getActiveWorksheet();
                for (var i = 1; i < range.text.length; i++) {

                    var sheet_row = i + 1;
                    var address = getCharFromNumber(header_then + 1) + sheet_row;

                    if (document.getElementById('if_operator').value == "equal") {
                        if (range.values[i][header_if] == ifcondition) {
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
                                if (range.values[i][header_then] < thencondition || range.values[i][header_then] > betweencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "notbetween") {
                                if (range.values[i][header_then] > thencondition && range.values[i][header_then] < betweencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "inlist") {
                                var check = 0;
                                for (var run = 0; run < splitted_then_list.length; run++) {
                                    if (range.values[i][header_then] == splitted_then_list[run]) {
                                        check = 1;
                                    }
                                }
                                if (check == 0){
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                        }
                    }

                    if (document.getElementById('if_operator').value == "smaller") {
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
                                if (range.values[i][header_then] < thencondition || range.values[i][header_then] > betweencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "notbetween") {
                                if (range.values[i][header_then] > thencondition && range.values[i][header_then] < betweencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "inlist") {
                                var check = 0;
                                for (var run = 0; run < splitted_then_list.length; run++) {
                                    if (range.values[i][header_then] == splitted_then_list[run]) {
                                        check = 1;
                                    }
                                }
                                if (check == 0){
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                        }
                    }

                    if (document.getElementById('if_operator').value == "greater") {
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
                                if (range.values[i][header_then] < thencondition || range.values[i][header_then] > betweencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "notbetween") {
                                if (range.values[i][header_then] > thencondition && range.values[i][header_then] < betweencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "inlist") {
                                var check = 0;
                                for (var run = 0; run < splitted_then_list.length; run++) {
                                    if (range.values[i][header_then] == splitted_then_list[run]) {
                                        check = 1;
                                    }
                                }
                                if (check == 0){
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                        }
                    }
                    if (document.getElementById('if_operator').value == "inequal") {
                        if (range.values[i][header_if] != ifcondition) {
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
                                if (range.values[i][header_then] < thencondition || range.values[i][header_then] > betweencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "notbetween") {
                                if (range.values[i][header_then] > thencondition && range.values[i][header_then] < betweencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "inlist") {
                                var check = 0;
                                for (var run = 0; run < splitted_then_list.length; run++) {
                                    if (range.values[i][header_then] == splitted_then_list[run]) {
                                        check = 1;
                                    }
                                }
                                if (check == 0){
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                        }
                    }

                    if (document.getElementById('if_operator').value == "between") {
                        if (range.values[i][header_if] > ifcondition && range.values[i][header_if] < ifbetweencondition) {
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
                                if (range.values[i][header_then] < thencondition || range.values[i][header_then] > betweencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "notbetween") {
                                if (range.values[i][header_then] > thencondition && range.values[i][header_then] < betweencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "inlist") {
                                var check = 0;
                                for (var run = 0; run < splitted_then_list.length; run++) {
                                    if (range.values[i][header_then] == splitted_then_list[run]) {
                                        check = 1;
                                    }
                                }
                                if (check == 0){
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                        }
                    }

                    if (document.getElementById('if_operator').value == "notbetween") {
                        if (range.values[i][header_if] < ifcondition || range.values[i][header_if] > ifbetweencondition) {
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
                                if (range.values[i][header_then] < thencondition || range.values[i][header_then] > betweencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "notbetween") {
                                if (range.values[i][header_then] > thencondition && range.values[i][header_then] < betweencondition) {
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                            else if (document.getElementById('then_operator').value == "inlist") {
                                var check = 0;
                                for (var run = 0; run < splitted_then_list.length; run++) {
                                    if (range.values[i][header_then] == splitted_then_list[run]) {
                                        check = 1;
                                    }
                                }
                                if (check == 0){
                                    highlightContentInWorksheet(act_worksheet, address, "red");
                                }
                            }
                        }
                    }

                    if (document.getElementById('if_operator').value == "inlist") {
                        for (var run = 0; run < splitted_if_list.length; run++) {
                            if (range.values[i][header_if] == splitted_if_list[run]) {
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
                                    if (range.values[i][header_then] < thencondition || range.values[i][header_then] > betweencondition) {
                                        highlightContentInWorksheet(act_worksheet, address, "red");
                                    }
                                }
                                else if (document.getElementById('then_operator').value == "notbetween") {
                                    if (range.values[i][header_then] > thencondition && range.values[i][header_then] < betweencondition) {
                                        highlightContentInWorksheet(act_worksheet, address, "red");
                                    }
                                }
                                else if (document.getElementById('then_operator').value == "inlist") {
                                    var check = 0;
                                    for (var runthen = 0; runthen < splitted_then_list.length; runthen++) {
                                        if (range.values[i][header_then] == splitted_then_list[run]) {
                                            check = 1;
                                        }
                                    }
                                    if (check == 0){
                                        highlightContentInWorksheet(act_worksheet, address, "red");
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