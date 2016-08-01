//redirect to detect inconsistencies
function redirectRule() {
    Office.context.document.settings.set('from_inconsistencies', false);
    window.location = "inconsistency.html";
}

//show textfield for ending delimiter if custom is selected
function displayBetween(){
    if(document.getElementById('then_operator').value == "between" || document.getElementById('then_operator').value == "notbetween") {
        document.getElementById('thenplaceholder').innerHTML = "Range start (included)";
        $('#betweenand').show();
        $('#explanationand').show();
    }
    else {
        $('#betweenand').hide();
        $('#explanationand').hide();
        if(document.getElementById('then_operator').value == "inlist") {
            document.getElementById('thenplaceholder').innerHTML = "Option1, Option2,...";
        }
        else {
            document.getElementById('thenplaceholder').innerHTML = "Type condition";
        }
    }
}

//get active textfield where to enter the selection of the user
var activeSelection = 0;
function setFocus(activeID) {
    activeSelection = activeID;
}

//display additional text field when between or not between operator is selected
function displaySimpleBetween(){
    if(document.getElementById('if_operator1').value == "between" || document.getElementById('if_operator1').value == "notbetween") {
        document.getElementById('ifplaceholder').innerHTML = "Range start (included)";
        $('#between_beginning1').show();
        $('#explanation_and').show();
    }
    else {
        $('#between_beginning1').hide();
        $('#explanation_and').hide();

        if(document.getElementById('if_operator1').value == "inlist") {
            document.getElementById('ifplaceholder').innerHTML = "Option1, Option2,...";
        }
        else {
            document.getElementById('ifplaceholder').innerHTML = "Type condition";
        }
    }
}

function showEnterpriseDialog() {
    document.getElementById('showEnterprise').style.visibility = 'visible';
}

(function () {
    // 'use strict';
    var count_drop = 1;

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            Office.context.document.settings.set('same_header_validation', false);
            Office.context.document.settings.set('last_clicked_function', "validation.html");
            if (Office.context.document.settings.get('prepjet_loaded_before') == null) {
                Office.context.document.settings.set('prepjet_loaded_before', true);
                Office.context.document.settings.saveAsync();
                window.location = "intro.html";
            }

            app.initialize();
            fillIfColumn();
            fillThenColumn();

            $('#between_beginning1').hide();
            $('#explanation_and').hide();
            $('#explanationand').hide();
            $('#betweenand').hide();
            $('#delimiter_beginning1').show();
            $('#apply_advanced').show();
            $('#to_inconsistency').hide();


            $(".dropdown_table").Dropdown();
            $(".ms-TextField").TextField();


            $('#and_cond').click(showEnterpriseDialog);
            $('#or_cond').click(showEnterpriseDialog);
            $('#apply_advanced').click(validationAndAdvanced);
            $('#to_inconsistency').click(redirectRule);
            $('#buttonOk').click(highlightHeader);


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
                document.getElementById('if_condition1').value = message;
                document.getElementById('ifplaceholder').style.visibility = 'hidden';
            }
            function writeifand(message){
                document.getElementById('if_between_condition1').value = message;
                document.getElementById('ifandplaceholder').style.visibility = 'hidden';
            }
            function writethen(message){
                document.getElementById('then_condition').value = message;
                document.getElementById('thenplaceholder').style.visibility = 'hidden';
            }
            function writethenand(message){
                document.getElementById('between_and').value = message;
                document.getElementById('thenandplaceholder').style.visibility = 'hidden';
            }


            // Hides error message dialogs when clicking on ok or close
            document.getElementById("buttonClose").onclick = function () {
                document.getElementById('showEmbeddedDialog').style.visibility = 'hidden';
            }
            document.getElementById("buttonCloseEnterprise").onclick = function () {
                document.getElementById('showEnterprise').style.visibility = 'hidden';
            }
            document.getElementById("buttonOkEnterprise").onclick = function () {
                document.getElementById('showEnterprise').style.visibility = 'hidden';
            }


            //show and hide help callout
            document.getElementById("help_icon").onclick = function () {
                document.getElementById('helpCallout').style.visibility = 'visible';
            }
            document.getElementById("closeCallout").onclick = function () {
                document.getElementById('helpCallout').style.visibility = 'hidden';
            }


            Excel.run(function (ctx) {

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
                                    window.location = "validation.html"
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
                    window.location = "validation.html";
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



    function highlightHeader() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('text');

            return ctx.sync().then(function() {

                for (var run = 0; run < range.text[0].length - 1; run++) {
                    for (var run2 = run + 1; run2 < range.text[0].length; run2++) {
                        if (range.text[0][run] == range.text[0][run2]) {
                            document.getElementById('showEmbeddedDialog').style.visibility = 'hidden';
                            highlightContentInWorksheet(worksheet, getCharFromNumber(run) + 1, '#EA7F04');
                            highlightContentInWorksheet(worksheet, getCharFromNumber(run2) + 1, '#EA7F04');
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


    //function to populate dropdown for if condition with column headers
    function fillIfColumn(){

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();
            range.load('text');

            return ctx.sync().then(function() {

                for (var run = 0; run < range.text[0].length - 1; run++) {
                    for (var run2 = run + 1; run2 < range.text[0].length; run2++) {
                        if (range.text[0][run] == range.text[0][run2]) {
                            document.getElementById('showEmbeddedDialog').style.visibility = 'visible';
                            Office.context.document.settings.set('same_header_validation', true);
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

                    document.getElementById("column_simple1").appendChild(el);
                }
                var cont_tmp = "table_simple1";
                $("." + cont_tmp).Dropdown();
                $("span.ms-Dropdown-title:empty").text(range.text[0][0]);
            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

    }

    //function to populate dropdowns with column headers
    function fillThenColumn(){

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();
            range.load('text');

            return ctx.sync().then(function() {

                for (var run = 0; run < range.text[0].length - 1; run++) {
                    for (var run2 = run + 1; run2 < range.text[0].length; run2++) {
                        if (range.text[0][run] == range.text[0][run2]) {
                            document.getElementById('showEmbeddedDialog').style.visibility = 'visible';
                            Office.context.document.settings.set('same_header_validation', true);
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

                    document.getElementById("column2_options").appendChild(el);
                }

                $(".dropdown_table_col2").Dropdown();
                $("span.ms-Dropdown-title:empty").text(range.text[0][0]);

            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

    }



    //validation when advanced rule is selected
    function validationAndAdvanced() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();
            var selected_identifier2 = document.getElementById('column2_options').value;


            if (document.getElementById('then_operator').value == "inlist") {
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

            range.load('text');

            return ctx.sync().then(function() {

                var act_worksheet = ctx.workbook.worksheets.getActiveWorksheet();
                var selected_identifier1 = document.getElementById('column_simple1').value;

                for (var k = 0; k < range.text[0].length; k++){
                    if (selected_identifier1 == range.text[0][k] || selected_identifier1 == "Column " + getCharFromNumber(k)){
                        var header_if = k;
                    }
                }
                for (var k = 0; k < range.text[0].length; k++){
                    if (selected_identifier2 == range.text[0][k] || selected_identifier2 == "Column " + getCharFromNumber(k)){
                        var header_then = k;
                    }
                }


                function highlightThenCond() {
                    var address = getCharFromNumber(header_then) + sheet_row;
                    if (document.getElementById('then_operator').value == "equal") {
                        if (range.text[i][header_then] != thencondition) {
                            highlightContentInWorksheet(act_worksheet, address, '#EA7F04');
                        }
                    }
                    else if (document.getElementById('then_operator').value == "smaller") {
                         if (range.text[i][header_then] >= thencondition) {
                            highlightContentInWorksheet(act_worksheet, address, '#EA7F04');
                         }
                    }
                    else if (document.getElementById('then_operator').value == "greater") {
                        if (range.text[i][header_then] <= thencondition) {
                            highlightContentInWorksheet(act_worksheet, address, '#EA7F04');
                        }
                    }
                    else if (document.getElementById('then_operator').value == "inequal") {
                        if (range.text[i][header_then] == thencondition) {
                            highlightContentInWorksheet(act_worksheet, address, '#EA7F04');
                        }
                    }
                    else if (document.getElementById('then_operator').value == "between") {
                        if (range.text[i][header_then] < thencondition || range.text[i][header_then] > betweencondition) {
                            highlightContentInWorksheet(act_worksheet, address, '#EA7F04');
                        }
                    }
                    else if (document.getElementById('then_operator').value == "notbetween") {
                        if (range.text[i][header_then] > thencondition && range.text[i][header_then] < betweencondition) {
                            highlightContentInWorksheet(act_worksheet, address, '#EA7F04');
                        }
                    }
                    else if (document.getElementById('then_operator').value == "inlist") {
                        var check_then = 0;
                        for (var runthen = 0; runthen < splitted_then_list.length; runthen++) {
                            if (range.text[i][header_then] == splitted_then_list[runthen]) {
                                check_then = 1;
                            }
                        }
                        if (check_then == 0){
                            highlightContentInWorksheet(act_worksheet, address, '#EA7F04');
                        }
                    }
                }


                //go through all rows and check if if condition is true
                for (var i = 1; i < range.text.length; i++) {
                    var sheet_row = i + 1;

                    if (document.getElementById('if_operator1').value == "inlist") {
                        var in_if_list = document.getElementById('if_condition1').value;
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
                        if (isNaN(Number(document.getElementById('if_condition1').value)) == true) {
                            var ifcondition = document.getElementById('if_condition1').value;
                        }
                        else {
                            var ifcondition = Number(document.getElementById('if_condition1').value);
                        }
                    }

                    //get correct value in if condition for between/not between 2nd value
                    if (document.getElementById('if_operator1').value == "between" || document.getElementById('if_operator1').value == "notbetween") {
                        if (isNaN(Number(document.getElementById('if_between_condition1').value)) == true) {
                            var ifbetweencondition = document.getElementById('if_between_condition1').value;
                        }
                        else {
                            var ifbetweencondition = Number(document.getElementById('if_between_condition1').value);
                        }
                    }

                    if (document.getElementById('if_operator1').value == "equal") {
                        if (range.text[i][header_if] == ifcondition) {
                            highlightThenCond();
                        }
                    }

                    if (document.getElementById('if_operator1').value == "smaller") {
                        if (range.text[i][header_if] < ifcondition) {
                            highlightThenCond();
                        }
                    }

                    if (document.getElementById('if_operator1').value == "greater") {
                        if (range.text[i][header_if] > ifcondition) {
                            highlightThenCond();
                        }
                    }

                    if (document.getElementById('if_operator1').value == "inequal") {
                        if (range.text[i][header_if] != ifcondition) {
                            highlightThenCond();
                        }
                    }

                    if (document.getElementById('if_operator1').value == "between") {
                        if (range.text[i][header_if] > ifcondition && range.text[i][header_if] < ifbetweencondition) {
                            highlightThenCond();
                        }
                    }

                    if (document.getElementById('if_operator1').value == "notbetween") {
                        if (range.text[i][header_if] < ifcondition || range.text[i][header_if] > ifbetweencondition) {
                            highlightThenCond();
                        }
                    }

                    if (document.getElementById('if_operator1').value == "inlist") {
                        var check_list = 0;
                        for (var run = 0; run < splitted_if_list.length; run++) {
                            if (range.text[i][header_if] == splitted_if_list[run]) {
                                check_list = 1;
                            }
                        }
                        if (check_list == 1) {
                            highlightThenCond();
                        }
                    }
                }

                window.location = "validation.html";
            });


        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

})();