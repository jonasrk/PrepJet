//redirect to detect inconsistencies
function redirectRule() {
    Office.context.document.settings.set('from_inconsistencies', false);
    window.location = "inconsistency.html";
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

//get active textfield where to enter the selection of the user
var activeSelection = 0;
function setFocus(activeID) {
    activeSelection = activeID;
}

//display additional text field when between or not between operator is selected
function displaySimpleBetween(k){
    if(document.getElementById('if_operator' + k).value == "between" || document.getElementById('if_operator' + k).value == "notbetween") {
        $('#between_beginning' + k).show();
    }
    else {
        $('#between_beginning' + k).hide();
    }
}

function showEnterpriseDialog() {
    document.getElementById('showEnterprise').style.visibility = 'visible';
}

(function () {
    'use strict';
    var count_drop = 1;
    var mixed_condition = [];
    mixed_condition.push(1);

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            Office.context.document.settings.set('same_header_validation', false);
            //Office.context.document.settings.set('last_condition_added', 'simple');
            //Office.context.document.settings.set('populated_then', false);
            Office.context.document.settings.set('last_clicked_function', "validation.html");
            if (Office.context.document.settings.get('prepjet_loaded_before') == null) {
                Office.context.document.settings.set('prepjet_loaded_before', true);
                Office.context.document.settings.saveAsync();
                window.location = "intro.html";
            }

            app.initialize();
            fillIfColumn();
            fillThenColumn();

            //$('#tmp_hide').hide();
            $('#between_beginning1').hide();
            $('#betweenand').hide();
            $('#delimiter_beginning1').show();
            //$('#remove_cond').hide();
            //$('#apply_mixed_simple').hide();
            $('#apply_advanced').show();
            //$('#apply_or_advanced').hide();
            //$('#apply_mixed_advanced').hide();
            //$('#apply_or_simple').hide();
            $('#to_inconsistency').hide();

            if (Office.context.document.settings.get('from_inconsistencies') == true){
                $('#to_inconsistency').show();
            }

            $(".dropdown_table").Dropdown();
            $(".ms-TextField").TextField();

            //$('#apply_and_simple').click(validationAndSimple);
            //$('#apply_or_simple').click(validationOrSimple);
            //$('#apply_mixed_simple').click(validationMixedSimple);
            $('#and_cond').click(showEnterpriseDialog);
            $('#or_cond').click(showEnterpriseDialog);
            //$('#then_cond').click(addThenCondition);
            //$('#remove_cond').click(removeCondition);
            $('#apply_advanced').click(validationAndAdvanced);
            //$('#apply_or_advanced').click(validationOrAdvanced);
            //$('#apply_mixed_advanced').click(validationMixedAdvanced);
            $('#to_inconsistency').click(redirectRule);


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
            }

            function writeifand(message){
                document.getElementById('if_between_condition1').value = message;
            }

            function writethen(message){
                document.getElementById('then_condition').value = message;
            }

            function writethenand(message){
                document.getElementById('between_and').value = message;
            }


            // Hides the dialog.
            document.getElementById("buttonClose").onclick = function () {
                $("#showEmbeddedDialog").hide();
            }

            // Performs the action and closes the dialog.
            document.getElementById("buttonOk").onclick = function () {
                $("#showEmbeddedDialog").hide();
            }

            // Hides the dialog.
            document.getElementById("buttonCloseEnterprise").onclick = function () {
                document.getElementById('showEnterprise').style.visibility = 'hidden';
                //$("#showEnterprise").hide();
            }

            // Performs the action and closes the dialog.
            document.getElementById("buttonOkEnterprise").onclick = function () {
                $("#showEnterprise").hide();
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
                            highlightContentInWorksheet(worksheet, getCharFromNumber(run) + 1, '#EA7F04');
                            highlightContentInWorksheet(worksheet, getCharFromNumber(run2) + 1, '#EA7F04');
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

                    document.getElementById("column_simple" + count_drop).appendChild(el);
                }
                var cont_tmp = "table_simple" + count_drop;
                console.log($("." + cont_tmp));
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
                            highlightContentInWorksheet(worksheet, getCharFromNumber(run) + 1, '#EA7F04');
                            highlightContentInWorksheet(worksheet, getCharFromNumber(run2) + 1, '#EA7F04');
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

                //var cont_tmp = "table_simple" + count_drop;
                //console.log($("." + cont_tmp));
                //$("." + cont_tmp).Dropdown();
                //$("span.ms-Dropdown-title:empty").text(range.text[0][0]);
            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

    }


    //function to check condition when only AND statements are used
    /*function validationAndSimple() {
        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            //get used range in active Sheet
            range.load('text');
            range.load('valueTypes');
            range.load('values');
            var range_all_adding_to = worksheet.getRange();
            var range_adding_to = range_all_adding_to.getUsedRange();
            range_adding_to.load('address');
            range_adding_to.load('text');


            return ctx.sync().then(function() {

                var header_if = [];
                var selected_identifier = [];
                for (var k = 0; k < count_drop; k++) {
                    selected_identifier.push(document.getElementById('column_simple' + (k + 1)).value);
                }

                for (var runsel = 0; runsel < selected_identifier.length; runsel++) {
                    for (var k = 0; k < range.text[0].length; k++){
                        if (selected_identifier[runsel] == range.text[0][k] || selected_identifier[runsel] == "Column " + getCharFromNumber(k)){
                            header_if.push(k);
                        }
                    }
                }

                for (var runcon = 0; runcon < count_drop; runcon++){

                    if (document.getElementById('if_operator' + (runcon + 1)).value == "inlist") {
                        var in_list = document.getElementById('if_condition' + (runcon + 1)).value;
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
                        if (isNaN(Number(document.getElementById('if_condition' + (runcon + 1)).value)) == true) {
                            var ifcondition = document.getElementById('if_condition' + (runcon + 1)).value;
                        }
                        else {
                            var ifcondition = Number(document.getElementById('if_condition' + (runcon + 1)).value);
                        }
                    }

                    if (document.getElementById('if_operator' + (runcon + 1)).value == "notbetween" || document.getElementById('if_operator' + (runcon + 1)).value == "between") {
                        if (isNaN(Number(document.getElementById('if_between_condition' + (runcon + 1)).value)) == true) {
                            var ifbetweencondition = document.getElementById('if_between_condition' + (runcon + 1)).value;
                        }
                        else {
                            var ifbetweencondition = Number(document.getElementById('if_between_condition' + (runcon + 1)).value);
                        }
                    }

                    var act_worksheet = ctx.workbook.worksheets.getActiveWorksheet();
                    var col_index = header_if[runcon];
                    for (var i = 1; i < range.text.length; i++) {

                        var check_cond = 0;
                        if (document.getElementById('if_operator' + (runcon + 1)).value == "equal") {
                            if (range.text[i][col_index] != ifcondition) {
                                check_cond += 1;
                            }
                        }

                        if (document.getElementById('if_operator' + (runcon + 1)).value == "smaller") {
                            if (range.text[i][col_index] >= ifcondition) {
                                check_cond += 1;
                            }
                        }

                        if (document.getElementById('if_operator' + (runcon + 1)).value == "greater") {
                            if (range.text[i][col_index] <= ifcondition) {
                                check_cond += 1;
                            }
                        }

                        if (document.getElementById('if_operator' + (runcon + 1)).value == "inequal") {
                            if (range.text[i][col_index] == ifcondition) {
                                check_cond += 1;
                            }
                        }

                        if (document.getElementById('if_operator' + (runcon + 1)).value == "between") {
                            if (range.text[i][col_index] < ifcondition || range.text[i][col_index] > ifbetweencondition) {
                                 check_cond += 1;
                            }
                        }

                        if (document.getElementById('if_operator' + (runcon + 1)).value == "notbetween") {
                            if (range.text[i][col_index] > ifcondition && range.text[i][col_index] < ifbetweencondition) {
                                 check_cond += 1;
                            }
                        }

                        if (document.getElementById('if_operator' + (runcon + 1)).value == "inlist") {
                            var check = 0;
                            for (run = 0; run < splitted_list.length; run++) {
                                if (range.text[i][col_index] == splitted_list[run]) {
                                     check = 1;
                                }
                            }
                            if (check != 1) {
                                check_cond += 1;
                            }
                        }

                        var sheet_row = i + 1;
                        if (check_cond > 0) {
                            for (var k = 0; k < header_if.length; k++) {
                                var address = getCharFromNumber(header_if[k]) + sheet_row;
                                highlightContentInWorksheet(act_worksheet, address, '#EA7F04');
                            }
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
    }*/

    //function to check condition  if only OR statements are used
    /*function validationOrSimple() {
        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            //get used range in active Sheet
            range.load('text');
            range.load('valueTypes');
            range.load('values');
            var range_all_adding_to = worksheet.getRange();
            var range_adding_to = range_all_adding_to.getUsedRange();
            range_adding_to.load('address');
            range_adding_to.load('text');


            return ctx.sync().then(function() {

                var header_if = [];
                var selected_identifier = [];
                for (var k = 0; k < count_drop; k++) {
                    selected_identifier.push(document.getElementById('column_simple' + (k + 1)).value);
                }

                for (var runsel = 0; runsel < selected_identifier.length; runsel++) {
                    for (var k = 0; k < range.text[0].length; k++){
                        if (selected_identifier[runsel] == range.text[0][k] || selected_identifier == "Column " + getCharFromNumber(k)){
                            header_if.push(k);
                        }
                    }
                }

                for (var i = 1; i < range.text.length; i++) {
                    var check_cond = 0;
                    for (var runcon = 0; runcon < count_drop; runcon++){

                        if (document.getElementById('if_operator' + (runcon + 1)).value == "inlist") {
                            var in_list = document.getElementById('if_condition' + (runcon + 1)).value;
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
                            if (isNaN(Number(document.getElementById('if_condition' + (runcon + 1)).value)) == true) {
                                var ifcondition = document.getElementById('if_condition' + (runcon + 1)).value;
                            }
                            else {
                                var ifcondition = Number(document.getElementById('if_condition' + (runcon + 1)).value);
                            }
                        }

                        if (document.getElementById('if_operator' + (runcon + 1)).value == "notbetween" || document.getElementById('if_operator' + (runcon + 1)).value == "between") {
                            if (isNaN(Number(document.getElementById('if_between_condition' + (runcon + 1)).value)) == true) {
                                var ifbetweencondition = document.getElementById('if_between_condition' + (runcon + 1)).value;
                            }
                            else {
                                var ifbetweencondition = Number(document.getElementById('if_between_condition' + (runcon + 1)).value);
                            }
                        }

                        //loop through whole column to extract value from
                        var act_worksheet = ctx.workbook.worksheets.getActiveWorksheet();
                        var col_index = header_if[runcon];

                        if (document.getElementById('if_operator' + (runcon + 1)).value == "equal") {
                            if (range.text[i][col_index] != ifcondition) {
                                check_cond += 1;
                            }
                        }

                        if (document.getElementById('if_operator' + (runcon + 1)).value == "smaller") {
                            if (range.text[i][col_index] >= ifcondition) {
                                check_cond += 1;
                            }
                        }

                        if (document.getElementById('if_operator' + (runcon + 1)).value == "greater") {
                            if (range.text[i][col_index] <= ifcondition) {
                                check_cond += 1;
                            }
                        }

                        if (document.getElementById('if_operator' + (runcon + 1)).value == "inequal") {
                            if (range.text[i][col_index] == ifcondition) {
                                check_cond += 1;
                            }
                        }

                        if (document.getElementById('if_operator' + (runcon + 1)).value == "between") {
                            if (range.text[i][col_index] < ifcondition || range.text[i][col_index] > ifbetweencondition) {
                                 check_cond += 1;
                            }
                        }

                        if (document.getElementById('if_operator' + (runcon + 1)).value == "notbetween") {
                            if (range.text[i][col_index] > ifcondition && range.text[i][col_index] < ifbetweencondition) {
                                 check_cond += 1;
                            }
                        }

                        if (document.getElementById('if_operator' + (runcon + 1)).value == "inlist") {
                            var check = 0;
                            for (run = 0; run < splitted_list.length; run++) {
                                if (range.text[i][col_index] == splitted_list[run]) {
                                     check = 1;
                                }
                            }
                            if (check != 1) {
                                check_cond += 1;
                            }
                        }
                    }

                    var sheet_row = i + 1;
                    if (check_cond == count_drop) {
                        for (var k = 0; k < header_if.length; k++) {
                            var address = getCharFromNumber(header_if[k]) + sheet_row;
                            highlightContentInWorksheet(act_worksheet, address, '#EA7F04');
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
    }*/


    /*function addAndCondition (start_var) {

        Office.context.document.settings.set('last_condition_added', 'and');
        count_drop += 1;
        var div_head = document.createElement("div");
        div_head.id = "subhead" + count_drop;
        document.getElementById("condition_holder").appendChild(div_head);

        var label = document.createElement("label");
        label.id = "head" + count_drop;
        label.className = "ms-Label";
        label.innerHTML = "AND";
        div_head.appendChild(label);

        addDropdown(count_drop);
        fillSimpleColumn();
        addOperator(count_drop);
        addTextField(count_drop);
        addBetweenField(count_drop);
        document.getElementById('if_operator' + count_drop).setAttribute("onchange",  "displaySimpleBetween(" + count_drop + ")");

        var cont_tmp = "dropdown_table" + count_drop;
        console.log($("." + cont_tmp));
        $("." + cont_tmp).Dropdown();

        $('#between_beginning' + count_drop).hide();
        $('#remove_cond').show();

        mixed_condition.push(1);

        var check_mix = 0;
        var test = 0;
        for (var k = 0; k < mixed_condition.length; k++) {
                test += mixed_condition[k];
                check_mix = test % mixed_condition.length;
        }
        if (check_mix == 0 && test == mixed_condition.length) {
            $('#apply_and_simple').show();
            $('#apply_mixed_simple').hide();
            $('#apply_or_simple').hide();
        }
        else {
            for (var k = 1; k < mixed_condition.length; k++) {
                if (mixed_condition[k] == 1) {
                    $('#apply_mixed_simple').show();
                    $('#apply_and_simple').hide();
                    $('#apply_or_simple').hide();
                    break;
                }
                else {
                    $('#apply_or_simple').show();
                    $('#apply_and_simple').hide();
                    $('#apply_mixed_simple').hide();
                }
            }
        }
    }*/


    /*function addORCondition (start_var) {

        Office.context.document.settings.set('last_condition_added', 'or');
        $('#apply_or_simple').show();
        $('#apply_and_simple').hide();

        count_drop += 1;
        var div_head = document.createElement("div");
        div_head.id = "subhead" + count_drop;
        document.getElementById("condition_holder").appendChild(div_head);

        var label = document.createElement("label");
        label.id = "head" + count_drop;
        label.className = "ms-Label";
        label.innerHTML = "OR";
        div_head.appendChild(label);

        addDropdown(count_drop);
        fillSimpleColumn();
        addOperator(count_drop);
        addTextField(count_drop);
        addBetweenField(count_drop);
        $('#between_beginning' + count_drop).hide();

        var cont_tmp = "dropdown_table" + count_drop;
        console.log($("." + cont_tmp));
        $("." + cont_tmp).Dropdown();
        document.getElementById('if_operator' + count_drop).setAttribute("onchange",  "displaySimpleBetween(" + count_drop + ")");
        $('#remove_cond').show();

        mixed_condition.push(2);

        var check_mix = 0;
        var test = 0;
        for (var k = 0; k < mixed_condition.length; k++) {
                test += mixed_condition[k];
                check_mix = test % mixed_condition.length;
        }
        if (check_mix == 0 && test == mixed_condition.length) {
            $('#apply_and_simple').show();
            $('#apply_mixed_simple').hide();
            $('#apply_or_simple').hide();
        }
        else {
            for (var k = 1; k < mixed_condition.length; k++) {
                if (mixed_condition[k] == 1) {
                    $('#apply_mixed_simple').show();
                    $('#apply_and_simple').hide();
                    $('#apply_or_simple').hide();
                    break;
                }
                else {
                    $('#apply_or_simple').show();
                    $('#apply_and_simple').hide();
                    $('#apply_mixed_simple').hide();
                }
            }
        }
    }*/


    /*function addThenCondition () {

        $('#tmp_hide').show();
        $('#remove_cond').show();
        $('#apply_and_simple').hide();
        $('#apply_or_simple').hide();
        $('#apply_mixed_simple').hide();
        $('#betweenand').hide();
        $('#and_cond').hide();
        $('#or_cond').hide();
        $('#then_cond').hide();
        $('#add_label').hide();


        var check_mix = 0;
        var test = 0;
        for (var k = 0; k < mixed_condition.length; k++) {
                test += mixed_condition[k];
                check_mix = test % mixed_condition.length;
        }
        if (check_mix == 0 && test == mixed_condition.length) {
            $('#apply_advanced').show();
            $('#apply_mixed_advanced').hide();
            $('#apply_or_advanced').hide();
        }
        else {
            for (var k = 1; k < mixed_condition.length; k++) {
                if (mixed_condition[k] == 1) {
                    $('#apply_mixed_advanced').show();
                    $('#apply_advanced').hide();
                    $('#apply_or_advanced').hide();
                    break;
                }
                else {
                    $('#apply_or_advanced').show();
                    $('#apply_advanced').hide();
                    $('#apply_mixed_advanced').hide();
                }
            }
        }


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
                            el.value = "Column " + getCharFromNumber(i);
                            el.textContent = "Column " + getCharFromNumber(i);
                        }

                    document.getElementById('column2_options').appendChild(el);
                }

                if (Office.context.document.settings.get('populated_then') == false) {
                    console.log($(".dropdown_table_col2"));
                    $(".dropdown_table_col2").Dropdown();
                    $("span.ms-Dropdown-title:empty").text(range.text[0][0]);
                }

                Office.context.document.settings.set('then_condition_pressed', true);
                Office.context.document.settings.set('populated_then', true);

            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

    }*/

    /*function removeCondition() {

        if (Office.context.document.settings.get('then_condition_pressed') == true) {
            $('#tmp_hide').hide();
            $('#apply_advanced').hide();
            $('#apply_or_advanced').hide();
            $('#apply_mixed_advanced').hide();
            $('#and_cond').show();
            $('#or_cond').show();
            $('#then_cond').show();
            $('#add_label').show();

            Office.context.document.settings.set('then_condition_pressed', false);

            var check_mix = 0;
            var test = 0;
            for (var k = 0; k < mixed_condition.length; k++) {
                    test += mixed_condition[k];
                    check_mix = test % mixed_condition.length;
            }
            if (check_mix == 0 && test == mixed_condition.length) {
                $('#apply_and_simple').show();
                $('#apply_mixed_simple').hide();
                $('#apply_or_simple').hide();
            }
            else {
                for (var k = 1; k < mixed_condition.length; k++) {
                    if (mixed_condition[k] == 1) {
                        $('#apply_mixed_simple').show();
                        $('#apply_and_simple').hide();
                        $('#apply_or_simple').hide();
                        break;
                    }
                    else {
                        $('#apply_or_simple').show();
                        $('#apply_and_simple').hide();
                        $('#apply_mixed_simple').hide();
                    }
                }
            }
        }
        else {
            mixed_condition.pop();

            if (count_drop > 1) {
                var parent = document.getElementById('condition_holder');
                var child = document.getElementById('condition' + count_drop);
                var child_head = document.getElementById('subhead' + count_drop);

                parent.removeChild(child_head);
                parent.removeChild(child);

            }
            count_drop -= 1;

            var check_mix = 0;
            var test = 0;
            for (var k = 0; k < mixed_condition.length; k++) {
                    test += mixed_condition[k];
                    check_mix = test % mixed_condition.length;
            }
            if (check_mix == 0 && test == mixed_condition.length) {
                $('#apply_and_simple').show();
                $('#apply_mixed_simple').hide();
                $('#apply_or_simple').hide();
            }
            else {
                for (var k = 1; k < mixed_condition.length; k++) {
                    if (mixed_condition[k] == 1) {
                        $('#apply_mixed_simple').show();
                        $('#apply_and_simple').hide();
                        $('#apply_or_simple').hide();
                        break;
                    }
                    else {
                        $('#apply_or_simple').show();
                        $('#apply_and_simple').hide();
                        $('#apply_mixed_simple').hide();
                    }
                }
            }
        }

        if (count_drop == 1) {
            $('#remove_cond').hide();
        }

    }*/


    //validation when advanced rule is selected
    function validationAndAdvanced() {
        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();
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

                var act_worksheet = ctx.workbook.worksheets.getActiveWorksheet();

                var selected_identifier1 = [];
                for (var k = 0; k < count_drop; k++) {
                    selected_identifier1.push(document.getElementById('column_simple' + (k + 1)).value);
                }

                var header_if = [];
                for (var runsel = 0; runsel < selected_identifier1.length; runsel++) {
                    for (var k = 0; k < range.text[0].length; k++){
                        if (selected_identifier1[runsel] == range.text[0][k] || selected_identifier1[runsel] == "Column " + getCharFromNumber(k)){
                            header_if.push(k);
                        }
                    }
                }

                //get column in header for which to check then condition
                var header_then = 0;
                for (var k = 0; k < range.text[0].length; k++){
                    if (selected_identifier2 == range.text[0][k] || selected_identifier2 == "Column " + getCharFromNumber(k)){
                        header_then = k;
                    }
                }

                for (var i = 1; i < range.text.length; i++) {
                    var check_cond = 0;
                    var sheet_row = i + 1;
                    for (var runcol = 0; runcol < count_drop; runcol++) {
                        var col_index = header_if[runcol];
                        if (document.getElementById('if_operator' + (runcol + 1)).value == "inlist") {
                            var in_if_list = document.getElementById('if_condition' + (runcol + 1)).value;
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
                            if (isNaN(Number(document.getElementById('if_condition' + (runcol + 1)).value)) == true) {
                                var ifcondition = document.getElementById('if_condition' + (runcol + 1)).value;
                            }
                            else {
                                var ifcondition = Number(document.getElementById('if_condition' + (runcol + 1)).value);
                            }
                        }

                        //get correct value in if condition for between/not between 2nd value
                        if (document.getElementById('if_operator' + (runcol + 1)).value == "between" || document.getElementById('if_operator' + (runcol + 1)).value == "notbetween") {
                            if (isNaN(Number(document.getElementById('if_between_condition' + (runcol + 1)).value)) == true) {
                                var ifbetweencondition = document.getElementById('if_between_condition' + (runcol + 1)).value;
                            }
                            else {
                                var ifbetweencondition = Number(document.getElementById('if_between_condition' + (runcol + 1)).value);
                            }
                        }

                        if (document.getElementById('if_operator' + (runcol + 1)).value == "equal") {
                            if (range.text[i][col_index] == ifcondition) {
                                check_cond += 1;
                            }
                        }

                        if (document.getElementById('if_operator' + (runcol + 1)).value == "smaller") {
                            if (range.text[i][col_index] < ifcondition) {
                                check_cond += 1;
                            }
                        }

                        if (document.getElementById('if_operator' + (runcol + 1)).value == "greater") {
                            if (range.text[i][col_index] > ifcondition) {
                                check_cond += 1;
                            }
                        }
                        if (document.getElementById('if_operator' + (runcol + 1)).value == "inequal") {
                            if (range.text[i][col_index] != ifcondition) {
                                check_cond += 1;
                            }
                        }

                        if (document.getElementById('if_operator' + (runcol + 1)).value == "between") {
                            if (range.text[i][col_index] > ifcondition && range.text[i][col_index] < ifbetweencondition) {
                                check_cond += 1;
                            }
                        }

                        if (document.getElementById('if_operator' + (runcol + 1)).value == "notbetween") {
                            if (range.text[i][col_index] < ifcondition || range.text[i][col_index] > ifbetweencondition) {
                                check_cond += 1;
                            }
                        }

                        if (document.getElementById('if_operator' + (runcol + 1)).value == "inlist") {
                            var check_list = 0;
                            for (var run = 0; run < splitted_if_list.length; run++) {
                                if (range.text[i][col_index] == splitted_if_list[run]) {
                                    check_list = 1;
                                }
                            }
                            if (check_list == 1) {
                                check_cond += 1;
                            }
                        }
                    }

                    var address = getCharFromNumber(header_then) + sheet_row;
                    if (check_cond == count_drop) {
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


    function validationOrAdvanced() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();
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

                var act_worksheet = ctx.workbook.worksheets.getActiveWorksheet();
                var selected_identifier1 = [];
                for (var k = 0; k < count_drop; k++) {
                    selected_identifier1.push(document.getElementById('column_simple' + (k + 1)).value);
                }

                var header_if = [];
                for (var runsel = 0; runsel < selected_identifier1.length; runsel++) {
                    for (var k = 0; k < range.text[0].length; k++){
                        if (selected_identifier1[runsel] == range.text[0][k] || selected_identifier1[runsel] == "Column " + getCharFromNumber(k)){
                            header_if.push(k);
                        }
                    }
                }

                //get column in header for which to check then condition
                var header_then = 0;
                for (var k = 0; k < range.text[0].length; k++){
                    if (selected_identifier2 == range.text[0][k] || selected_identifier2 == "Column " + getCharFromNumber(k)){
                        header_then = k;
                    }
                }

                //get correct list with values entered by user
                for (var runcol = 0; runcol < count_drop; runcol++) {
                    var col_index = header_if[runcol];
                    if (document.getElementById('if_operator' + (runcol + 1)).value == "inlist") {
                        var in_if_list = document.getElementById('if_condition' + (runcol + 1)).value;
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
                        if (isNaN(Number(document.getElementById('if_condition' + (runcol + 1)).value)) == true) {
                            var ifcondition = document.getElementById('if_condition' + (runcol + 1)).value;
                        }
                        else {
                            var ifcondition = Number(document.getElementById('if_condition' + (runcol + 1)).value);
                        }
                    }

                    //get correct value in if condition for between/not between 2nd value
                    if (document.getElementById('if_operator' + (runcol + 1)).value == "between" || document.getElementById('if_operator' + (runcol + 1)).value == "notbetween") {
                        if (isNaN(Number(document.getElementById('if_between_condition' + (runcol + 1)).value)) == true) {
                            var ifbetweencondition = document.getElementById('if_between_condition' + (runcol + 1)).value;
                        }
                        else {
                            var ifbetweencondition = Number(document.getElementById('if_between_condition' + (runcol + 1)).value);
                        }
                    }

                    for (var i = 1; i < range.text.length; i++) {

                        var sheet_row = i + 1;
                        var address = getCharFromNumber(header_then) + sheet_row;

                        if (document.getElementById('if_operator' + (runcol + 1)).value == "equal") {
                            if (range.text[i][col_index] == ifcondition) {
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
                                    var check = 0;
                                    for (var run = 0; run < splitted_then_list.length; run++) {
                                        if (range.text[i][header_then] == splitted_then_list[run]) {
                                            check = 1;
                                        }
                                    }
                                    if (check == 0){
                                        highlightContentInWorksheet(act_worksheet, address, '#EA7F04');
                                    }
                                }
                            }
                        }

                        if (document.getElementById('if_operator' + (runcol + 1)).value == "smaller") {
                            if (range.text[i][col_index] < ifcondition) {
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
                                    var check = 0;
                                    for (var run = 0; run < splitted_then_list.length; run++) {
                                        if (range.text[i][header_then] == splitted_then_list[run]) {
                                            check = 1;
                                        }
                                    }
                                    if (check == 0){
                                        highlightContentInWorksheet(act_worksheet, address, '#EA7F04');
                                    }
                                }
                            }
                        }

                        if (document.getElementById('if_operator' + (runcol + 1)).value == "greater") {
                            if (range.text[i][col_index] > ifcondition) {
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
                                    var check = 0;
                                    for (var run = 0; run < splitted_then_list.length; run++) {
                                        if (range.text[i][header_then] == splitted_then_list[run]) {
                                            check = 1;
                                        }
                                    }
                                    if (check == 0){
                                        highlightContentInWorksheet(act_worksheet, address, '#EA7F04');
                                    }
                                }
                            }
                        }
                        if (document.getElementById('if_operator' + (runcol + 1)).value == "inequal") {
                            if (range.text[i][col_index] != ifcondition) {
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
                                    var check = 0;
                                    for (var run = 0; run < splitted_then_list.length; run++) {
                                        if (range.text[i][header_then] == splitted_then_list[run]) {
                                            check = 1;
                                        }
                                    }
                                    if (check == 0){
                                        highlightContentInWorksheet(act_worksheet, address, '#EA7F04');
                                    }
                                }
                            }
                        }

                        if (document.getElementById('if_operator' + (runcol + 1)).value == "between") {
                            if (range.text[i][col_index] > ifcondition && range.text[i][header_if] < ifbetweencondition) {
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
                                    var check = 0;
                                    for (var run = 0; run < splitted_then_list.length; run++) {
                                        if (range.text[i][header_then] == splitted_then_list[run]) {
                                            check = 1;
                                        }
                                    }
                                    if (check == 0){
                                        highlightContentInWorksheet(act_worksheet, address, '#EA7F04');
                                    }
                                }
                            }
                        }

                        if (document.getElementById('if_operator' + (runcol + 1)).value == "notbetween") {
                            if (range.text[i][col_index] < ifcondition || range.text[i][header_if] > ifbetweencondition) {
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
                                    var check = 0;
                                    for (var run = 0; run < splitted_then_list.length; run++) {
                                        if (range.text[i][header_then] == splitted_then_list[run]) {
                                            check = 1;
                                        }
                                    }
                                    if (check == 0){
                                        highlightContentInWorksheet(act_worksheet, address, '#EA7F04');
                                    }
                                }
                            }
                        }

                        if (document.getElementById('if_operator' + (runcol + 1)).value == "inlist") {
                            for (var run = 0; run < splitted_if_list.length; run++) {
                                if (range.text[i][col_index] == splitted_if_list[run]) {
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
                                        var check = 0;
                                        for (var runthen = 0; runthen < splitted_then_list.length; runthen++) {
                                            if (range.text[i][header_then] == splitted_then_list[run]) {
                                                check = 1;
                                            }
                                        }
                                        if (check == 0){
                                            highlightContentInWorksheet(act_worksheet, address, '#EA7F04');
                                        }
                                    }
                                }
                            }
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


    function validationMixedSimple() {
        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            //get used range in active Sheet
            range.load('text');
            range.load('valueTypes');
            range.load('values');
            var range_all_adding_to = worksheet.getRange();
            var range_adding_to = range_all_adding_to.getUsedRange();
            range_adding_to.load('address');
            range_adding_to.load('text');


            return ctx.sync().then(function() {

                var act_worksheet = ctx.workbook.worksheets.getActiveWorksheet();
                var or_index = [];
                for (var k = 0; k < mixed_condition; k++) {
                    if (mixed_condition[k] == 2) {
                        or_index.push(k);
                    }
                }

                var and_index = [];
                var tmp = [0];
                for (var k = 1; k < mixed_condition.length; k++) {
                    if (mixed_condition[k] == 2) {
                        and_index.push(tmp);
                        tmp = [k];
                    }
                    else {
                        tmp.push(k);
                    }
                }
                and_index.push(tmp);

                var selected_identifier = [];
                for (var i = 0; i < and_index.length; i++) {
                    var id_tmp = [];
                    for (var k = 0; k < and_index[i].length; k++) {
                        var col_ind = and_index[i][k];
                        id_tmp.push(document.getElementById('column_simple' + (col_ind + 1)).value);
                    }
                    selected_identifier.push(id_tmp);
                }

                var header_columns = [];
                for (var i = 0; i < selected_identifier.length;i++) {
                    var head_tmp = [];
                    for (var k = 0; k < selected_identifier[i].length; k++){
                        for (var j = 0; j < range.text[0].length; j++) {
                            if (selected_identifier[i][k] == range.text[0][j] || selected_identifier == "Column " + getCharFromNumber(j)){
                                head_tmp.push(j);
                            }
                        }
                    }
                    header_columns.push(head_tmp);
                }


                for (var i = 1; i < range.text.length; i++) {

                    var total_cond = 0;
                    for (var runcon = 0; runcon < header_columns.length; runcon++){

                        var check_cond = 0;
                        for (var run = 0; run < header_columns[runcon].length; run++){

                            if (document.getElementById('if_operator' + (and_index[runcon][run] + 1)).value == "inlist") {
                                var in_list = document.getElementById('if_condition' + (and_index[runcon][run] + 1)).value;
                                var splitted_list = in_list.split(",");
                                for (var k = 0; k < splitted_list.length; k++) {
                                    splitted_list[k] = splitted_list[k].trim();
                                }
                                for (var k = 0; k < splitted_list.length; k++) {
                                    if (isNaN(Number(splitted_list[k])) != true) {
                                        splitted_list[k] = Number(splitted_list[k]);
                                    }
                                }
                            }
                            else {
                                if (isNaN(Number(document.getElementById('if_condition' + (and_index[runcon][run] + 1)).value)) == true) {
                                    var ifcondition = document.getElementById('if_condition' + (and_index[runcon][run] + 1)).value;
                                }
                                else {
                                    var ifcondition = Number(document.getElementById('if_condition' + (and_index[runcon][run] + 1)).value);
                                }
                            }

                            if (document.getElementById('if_operator' + (and_index[runcon][run] + 1)).value == "notbetween" || document.getElementById('if_operator' + (and_index[runcon][run] + 1)).value == "between") {
                                if (isNaN(Number(document.getElementById('if_between_condition' + (and_index[runcon][run] + 1)).value)) == true) {
                                    var ifbetweencondition = document.getElementById('if_between_condition' + (and_index[runcon][run] + 1)).value;
                                }
                                else {
                                    var ifbetweencondition = Number(document.getElementById('if_between_condition' + (and_index[runcon][run] + 1)).value);
                                }
                            }

                            var col_index = header_columns[runcon][run];
                            if (document.getElementById('if_operator' + (and_index[runcon][run] + 1)).value == "equal") {
                                if (range.text[i][col_index] != ifcondition) {
                                    check_cond += 1;
                                }
                            }

                            if (document.getElementById('if_operator' + (and_index[runcon][run] + 1)).value == "smaller") {
                                if (range.text[i][col_index] >= ifcondition) {
                                    check_cond += 1;
                                }
                            }

                            if (document.getElementById('if_operator' + (and_index[runcon][run] + 1)).value == "greater") {
                                if (range.text[i][col_index] <= ifcondition) {
                                    check_cond += 1;
                                }
                            }

                            if (document.getElementById('if_operator' + (and_index[runcon][run] + 1)).value == "inequal") {
                                if (range.text[i][col_index] == ifcondition) {
                                    check_cond += 1;
                                }
                            }

                            if (document.getElementById('if_operator' + (and_index[runcon][run] + 1)).value == "between") {
                                if (range.text[i][col_index] < ifcondition || range.text[i][col_index] > ifbetweencondition) {
                                     check_cond += 1;
                                }
                            }

                            if (document.getElementById('if_operator' + (and_index[runcon][run] + 1)).value == "notbetween") {
                                if (range.text[i][col_index] > ifcondition && range.text[i][col_index] < ifbetweencondition) {
                                     check_cond += 1;
                                }
                            }

                            if (document.getElementById('if_operator' + (and_index[runcon][run] + 1)).value == "inlist") {
                                var check = 0;
                                for (var k = 0; k < splitted_list.length; k++) {
                                    if (range.text[i][col_index] == splitted_list[k]) {
                                         check = 1;
                                    }
                                }
                                if (check != 1) {
                                    check_cond += 1;
                                }
                            }
                        }

                        if (check_cond > 0) {
                            total_cond += 1;
                        }
                    }

                    if (total_cond >= header_columns.length) {
                        var sheet_row = i + 1;
                        for (var k = 0; k < header_columns.length; k++) {
                            for (var j = 0; j < header_columns[k].length; j++){
                                var address = getCharFromNumber(header_columns[k][j]) + sheet_row;
                                highlightContentInWorksheet(act_worksheet, address, '#EA7F04');
                            }
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


    function validationMixedAdvanced() {
        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

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

                var act_worksheet = ctx.workbook.worksheets.getActiveWorksheet();
                var or_index = [];
                for (var k = 0; k < mixed_condition; k++) {
                    if (mixed_condition[k] == 2) {
                        or_index.push(k);
                    }
                }

                var and_index = [];
                var tmp = [0];
                for (var k = 1; k < mixed_condition.length; k++) {
                    if (mixed_condition[k] == 2) {
                        and_index.push(tmp);
                        tmp = [k];
                    }
                    else {
                        tmp.push(k);
                    }
                }
                and_index.push(tmp);

                var selected_identifier = [];
                for (var i = 0; i < and_index.length; i++) {
                    var id_tmp = [];
                    for (var k = 0; k < and_index[i].length; k++) {
                        var col_ind = and_index[i][k];
                        id_tmp.push(document.getElementById('column_simple' + (col_ind + 1)).value);
                    }
                    selected_identifier.push(id_tmp);
                }

                var header_columns = [];
                for (var i = 0; i < selected_identifier.length;i++) {
                    var head_tmp = [];
                    for (var k = 0; k < selected_identifier[i].length; k++){
                        for (var j = 0; j < range.text[0].length; j++) {
                            if (selected_identifier[i][k] == range.text[0][j] || selected_identifier == "Column " + getCharFromNumber(j)){
                                head_tmp.push(j);
                            }
                        }
                    }
                    header_columns.push(head_tmp);
                }

                //get column in header for which to check then condition
                var header_then = 0;
                for (var k = 0; k < range.text[0].length; k++){
                    if (selected_identifier2 == range.text[0][k] || selected_identifier2 == "Column " + getCharFromNumber(k)){
                        header_then = k;
                    }
                }


                for (var i = 1; i < range.text.length; i++) {

                    var sum_cond = 0;
                    for (var runcon = 0; runcon < header_columns.length; runcon++){

                        var check_cond = 0;
                        for (var run = 0; run < header_columns[runcon].length; run++){

                            if (document.getElementById('if_operator' + (and_index[runcon][run] + 1)).value == "inlist") {
                                var in_list = document.getElementById('if_condition' + (and_index[runcon][run] + 1)).value;
                                var splitted_list = in_list.split(",");
                                for (var k = 0; k < splitted_list.length; k++) {
                                    splitted_list[k] = splitted_list[k].trim();
                                }
                                for (var k = 0; k < splitted_list.length; k++) {
                                    if (isNaN(Number(splitted_list[k])) != true) {
                                        splitted_list[k] = Number(splitted_list[k]);
                                    }
                                }
                            }
                            else {
                                if (isNaN(Number(document.getElementById('if_condition' + (and_index[runcon][run] + 1)).value)) == true) {
                                    var ifcondition = document.getElementById('if_condition' + (and_index[runcon][run] + 1)).value;
                                }
                                else {
                                    var ifcondition = Number(document.getElementById('if_condition' + (and_index[runcon][run] + 1)).value);
                                }
                            }

                            if (document.getElementById('if_operator' + (and_index[runcon][run] + 1)).value == "notbetween" || document.getElementById('if_operator' + (and_index[runcon][run] + 1)).value == "between") {
                                if (isNaN(Number(document.getElementById('if_between_condition' + (and_index[runcon][run] + 1)).value)) == true) {
                                    var ifbetweencondition = document.getElementById('if_between_condition' + (and_index[runcon][run] + 1)).value;
                                }
                                else {
                                    var ifbetweencondition = Number(document.getElementById('if_between_condition' + (and_index[runcon][run] + 1)).value);
                                }
                            }

                            var col_index = header_columns[runcon][run];
                            if (document.getElementById('if_operator' + (and_index[runcon][run] + 1)).value == "equal") {
                                if (range.text[i][col_index] == ifcondition) {
                                    check_cond += 1;
                                }
                            }

                            if (document.getElementById('if_operator' + (and_index[runcon][run] + 1)).value == "smaller") {
                                if (range.text[i][col_index] < ifcondition) {
                                    check_cond += 1;
                                }
                            }

                            if (document.getElementById('if_operator' + (and_index[runcon][run] + 1)).value == "greater") {
                                if (range.text[i][col_index] > ifcondition) {
                                    check_cond += 1;
                                }
                            }

                            if (document.getElementById('if_operator' + (and_index[runcon][run] + 1)).value == "inequal") {
                                if (range.text[i][col_index] != ifcondition) {
                                    check_cond += 1;
                                }
                            }

                            if (document.getElementById('if_operator' + (and_index[runcon][run] + 1)).value == "between") {
                                if (range.text[i][col_index] > ifcondition && range.text[i][col_index] < ifbetweencondition) {
                                     check_cond += 1;
                                }
                            }

                            if (document.getElementById('if_operator' + (and_index[runcon][run] + 1)).value == "notbetween") {
                                if (range.text[i][col_index] < ifcondition || range.text[i][col_index] > ifbetweencondition) {
                                     check_cond += 1;
                                }
                            }

                            if (document.getElementById('if_operator' + (and_index[runcon][run] + 1)).value == "inlist") {
                                var check = 0;
                                for (var k = 0; k < splitted_list.length; k++) {
                                    if (range.text[i][col_index] == splitted_list[k]) {
                                         check = 1;
                                    }
                                }
                                if (check != 1) {
                                    check_cond += 1;
                                }
                            }
                        }

                        if (check_cond >= header_columns[runcon].length) {
                            sum_cond += 1;
                        }
                    }
                    if (sum_cond > 0){
                        var sheet_row = i + 1;
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
                            var check = 0;
                            for (var run = 0; run < splitted_then_list.length; run++) {
                                if (range.text[i][header_then] == splitted_then_list[run]) {
                                    check = 1;
                                }
                            }
                            if (check == 0){
                                highlightContentInWorksheet(act_worksheet, address, '#EA7F04');
                            }
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