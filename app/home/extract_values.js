//show textfield for beginning delimiter if custom is selected
function displayFieldBegin(){
    if (document.getElementById('beginning_options').value == "custom_b"){
        $('#delimiter_beginning').show();
    }
    else {
        $('#delimiter_beginning').hide();
    }
}

//show textfield for ending delimiter if custom is selected
function displayFieldEnd(){
    if(document.getElementById('ending_options').value == "custom_e") {
        $('#delimiter_end').show();
    }
    else {
        $('#delimiter_end').hide();
    }
}

function displayAdvancedCount() {
        $('#del_count_start').show();
        $('.del_count_dropdown_s').show();
        $('#del_count_end').show();
        $('.del_count_dropdown_e').show();
        $('#advanced_settings').hide();
        $('#advanced_hide').show();
        Office.context.document.settings.set('more_option_extract', true);
    }

function hideAdvancedCount() {
        $('#del_count_start').hide();
        $('.del_count_dropdown_s').hide();
        $('#del_count_end').hide();
        $('.del_count_dropdown_e').hide();
        $('#advanced_settings').show();
        $('#advanced_hide').hide();
        Office.context.document.settings.set('more_option', false);
}


(function () {
    // 'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            Office.context.document.settings.set('same_header_extract', false);
            Office.context.document.settings.set('more_option', false);
            Office.context.document.settings.set('last_clicked_function', "extract_values.html");
            if (Office.context.document.settings.get('prepjet_loaded_before') == null) {
                Office.context.document.settings.set('prepjet_loaded_before', true);
                Office.context.document.settings.saveAsync();
                window.location = "intro.html";
            }

            app.initialize();
            fillColumn();


            $('#delimiter_end').hide();
            $('#delimiter_beginning').hide();
            $('#del_count_start').hide();
            $('.del_count_dropdown_s').Dropdown();
            $('.del_count_dropdown_s').hide();
            $('#del_count_end').hide();
            $('.del_count_dropdown_e').Dropdown();
            $('.del_count_dropdown_e').hide();
            $('#advanced_settings').show();
            $('#advanced_hide').hide();

            $(".dropdown_table").Dropdown();
            $(".ms-TextField").TextField();

            $('#extract_Value').click(extractValue);
            $('#advanced_settings').click(displayAdvancedCount);
            $('#advanced_hide').click(hideAdvancedCount);
            $('#buttonOk').click(highlightHeader);


            /*Office.context.document.addHandlerAsync("documentSelectionChanged", myHandler, function(result){}
            );

            // Event handler function.
            function myHandler(eventArgs){
                Excel.run(function (ctx) {
                    var selectedRange = ctx.workbook.getSelectedRange();
                    selectedRange.load('address');
                    return ctx.sync().then(function () {
                        write(selectedRange.address);
                    });
                });
            }

            // Function that writes to a div with id='message' on the page.
            function write(message){
                document.getElementById('target_column_input').value = message;
            }*/


            //Show and hide error message if columns have same header name
            document.getElementById("buttonClose").onclick = function () {
                document.getElementById('showEmbeddedDialog').style.visibility = 'hidden';
            }



            //show and hide help callout
            document.getElementById("help_icon").onclick = function () {
                document.getElementById('helpCallout').style.visibility = 'visible';
            }
            document.getElementById("closeCallout").onclick = function () {
                document.getElementById('helpCallout').style.visibility = 'hidden';
            }

            document.getElementById("refresh_icon").onclick = function () {
                window.location = "extract_values.html";
            }

            //hide result message
            document.getElementById("resultClose").onclick = function () {
                document.getElementById('resultDialog').style.visibility = 'hidden';
                window.location = "extract_values.html";
            }
            document.getElementById("resultOk").onclick = function () {
                document.getElementById('resultDialog').style.visibility = 'hidden';
                window.location = "extract_values.html";
            }


            /*Excel.run(function (ctx) {

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
                                    window.location = "extract_values.html"
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
                    window.location = "extract_values.html";
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
*/
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
                        if (range.text[0][run] == range.text[0][run2] && range.text[0][run] != "") {
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


    function fillColumn(){

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('text');
            return ctx.sync().then(function() {

                for (var run = 0; run < range.text[0].length - 1; run++) {
                    for (var run2 = run + 1; run2 < range.text[0].length; run2++) {
                        if (range.text[0][run] == range.text[0][run2] && range.text[0][run] != "") {
                            document.getElementById('showEmbeddedDialog').style.visibility = 'visible';
                            Office.context.document.settings.set('same_header_extract', true);
                        }
                    }
                }

                var sp = document.createElement("span");
                sp.innerHTML = range.text[0][0];
                sp.className = "ms-Dropdown-title";
                document.getElementById("column1_options").appendChild(sp);

                for (var i = 0; i < range.text[0].length; i++) {
                    var el = document.createElement("option");
                    if (range.text[0][i] != "") {
                        el.value = range.text[0][i];
                        el.textContent = range.text[0][i];
                    }
                    else {
                        el.value = "Column " + getCharFromNumber(i);
                        el.textContent = "Column " + getCharFromNumber(i - 1);
                    }
                    document.getElementById("column1_options").appendChild(el);
                }

                $(".dropdown_table_col").Dropdown();
                $("span.ms-Dropdown-title:empty").text(range.text[0][0]);
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
            if (split_beginning == "semicolon_b") {
                split_beginning = ";";
            }
            if (split_beginning == "comma_b") {
                split_beginning = ",";
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
            if (split_end == "semicolon_e") {
                split_end = ";";
            }
            if (split_end == "comma_e") {
                split_end = ",";
            }


            //if advanced settings are selected, get values for delimiter count
            if (Office.context.document.settings.get('more_option_extract') == true) {
                var count_delimiter_start = 0;
                if (document.getElementById('delimiter_count_start').value == "one") { count_delimiter_start = 1; }
                else if (document.getElementById('delimiter_count_start').value == "two") { count_delimiter_start = 2; }
                else if (document.getElementById('delimiter_count_start').value == "three") { count_delimiter_start = 3; }
                else if (document.getElementById('delimiter_count_start').value == "four") { count_delimiter_start = 4; }
                else if (document.getElementById('delimiter_count_start').value == "five") { count_delimiter_start = 5; }
                else if (document.getElementById('delimiter_count_start').value == "six") { count_delimiter_start = 6; }
                else if (document.getElementById('delimiter_count_start').value == "seven") { count_delimiter_start = 7; }
                else if (document.getElementById('delimiter_count_start').value == "eight") { count_delimiter_start = 8; }
                else if (document.getElementById('delimiter_count_start').value == "nine") { count_delimiter_start = 9; }

                var count_direction_start = document.getElementById('del_count_drop_start').value;

                var count_delimiter_end = 0;
                if (document.getElementById('delimiter_count_end').value == "one") { count_delimiter_end = 1; }
                else if(document.getElementById('delimiter_count_end').value == "two") { count_delimiter_end = 2; }
                else if(document.getElementById('delimiter_count_end').value == "three") { count_delimiter_end = 3; }
                else if(document.getElementById('delimiter_count_end').value == "four") { count_delimiter_end = 4; }
                else if(document.getElementById('delimiter_count_end').value == "five") { count_delimiter_end = 5; }
                else if(document.getElementById('delimiter_count_end').value == "six") { count_delimiter_end = 6; }
                else if(document.getElementById('delimiter_count_end').value == "seven") { count_delimiter_end = 7; }
                else if(document.getElementById('delimiter_count_end').value == "eight") { count_delimiter_end = 8; }
                else if(document.getElementById('delimiter_count_end').value == "nine") { count_delimiter_end = 9; }
                else if(document.getElementById('delimiter_count_end').value == "none") { count_delimiter_end = 0; }

                var count_direction_end = document.getElementById('del_count_drop_end').value;
            }
            else {
                var count_delimiter_end = 0;
            }

            //get used range in active Sheet
            range.load('text');
            var range_all_adding_to = worksheet.getRange();
            var range_adding_to = range_all_adding_to.getUsedRange();

            range_adding_to.load('address');
            range_adding_to.load('text');

            return ctx.sync().then(function() {

                backupForUndo(range_adding_to);

                var header = 0;
                //get column in header from which to extract value
                for (var k = 0; k < range.text[0].length; k++){
                    if (selected_identifier == range.text[0][k] || selected_identifier == "Column " + getCharFromNumber(k)){
                        header = k;
                    }
                }

                //insert empty cell into header column
                var act_worksheet = ctx.workbook.worksheets.getActiveWorksheet();
                var extract_count = 0;
                var empty_count = 0;
                var extracted_array = [];

                //loop through whole column to extract value from
                for (var i = 0; i < range.text.length; i++) {

                    //get index where to start extracting value
                    if (split_beginning == "col_beginning"){
                        var position1 = 0;
                    }
                    else {
                        if (count_delimiter_start){
                            var tmp_array = range.text[i][header].split(split_beginning);
                            if (count_direction_start == "right") {
                                var loop_end = tmp_array.length - count_delimiter_start;
                                var str1_tmp = tmp_array[0];
                                for (var k = 1; k < loop_end; k++) {
                                    str1_tmp = str1_tmp.concat(split_beginning, tmp_array[k]);
                                }
                                if (document.getElementById('demo-checkbox-unselected').checked == true) {
                                    var position1 = str1_tmp.length;
                                }
                                else {
                                    var position1 = str1_tmp.length + 1;
                                }
                            }
                            else {
                                var str1_tmp = tmp_array[0];
                                for (var k = 1; k < count_delimiter_start; k++) {
                                    str1_tmp = str1_tmp.concat(split_beginning, tmp_array[k]);
                                }
                                if (document.getElementById('demo-checkbox-unselected').checked == true) {
                                    var position1 = str1_tmp.length;
                                }
                                else {
                                    var position1 = str1_tmp.length + 1;
                                }
                            }
                        }
                        else {
                            var position1 = range.text[i][header].indexOf(split_beginning);
                            if (position1 == -1) {
                                position1 = range.text[i][header].length;
                            }
                            else if (document.getElementById('demo-checkbox-unselected').checked == false) {
                                position1 = range.text[i][header].indexOf(split_beginning) + 1
                            }
                        }
                    }

                    //get index where to end extracting value
                    if (split_end == "col_end") {
                        var position2 = range.text[i][header].length;
                    }
                    else {
                        if (range.text[i][header].indexOf(split_end) == -1) {
                            var position2 = 0;
                        }
                        else {
                            //when delimiter to start and end is different
                            if (split_beginning != split_end) {
                                if (count_delimiter_end != 0){
                                    var tmp_array = range.text[i][header].split(split_end);
                                    if (count_direction_end == "right") {
                                        var loop_end = tmp_array.length - count_delimiter_end;
                                        var str1_tmp = tmp_array[0];
                                        for (var k = 1; k < loop_end; k++) {
                                            str1_tmp = str1_tmp.concat(split_end, tmp_array[k]);
                                        }
                                    }
                                    else {
                                        var str1_tmp = tmp_array[0];
                                        for (var k = 1; k < count_delimiter_end; k++) {
                                            str1_tmp = str1_tmp.concat(split_end, tmp_array[k]);
                                        }
                                    }

                                    if (document.getElementById('demo-checkbox-unselected').checked == true) {
                                        var position2 = str1_tmp.length + 1;
                                    }
                                    else {
                                        var position2 = str1_tmp.length;
                                    }
                                }

                                if (count_delimiter_end == 0) {
                                    var position2 = range.text[i][header].indexOf(split_end);
                                    if (position2 == -1) {
                                        position2 = 0;
                                    }
                                    else if(document.getElementById('demo-checkbox-unselected').checked == true) {
                                        position2 = position2 + 1;
                                    }
                                }
                            }
                            else {
                            //when delimiter to start and end is the same
                                if(count_delimiter_end == 0 && document.getElementById('demo-checkbox-unselected').checked == true) {
                                    var tmp = range.text[i][header].substring(position1 + 1, range.text[i][header].length);
                                    var position2 = tmp.indexOf(split_end) + position1 + 2;
                                }
                                else if (count_delimiter_end == 0 && document.getElementById('demo-checkbox-unselected').checked == false){
                                    var tmp = range.text[i][header].substring(position1, range.text[i][header].length);
                                    var position2 = tmp.indexOf(split_end) + position1;
                                }
                                else {
                                    var tmp_array = range.text[i][header].split(split_end);
                                    var str2_tmp = tmp_array[0];
                                    if (count_direction_end == "left") {
                                        for (var k = 1; k < count_delimiter_end; k++) {
                                            str2_tmp = str2_tmp.concat(split_end, tmp_array[k]);
                                        }
                                    }
                                    else {
                                        var loop_end = tmp_array.length - count_delimiter_end;
                                        for (var k = 1; k < loop_end; k++) {
                                            str2_tmp = str2_tmp.concat(split_end, tmp_array[k]);
                                        }
                                    }
                                    if (document.getElementById('demo-checkbox-unselected').checked == true) {
                                        var position2 = str2_tmp.length + 1;
                                    }
                                    else {
                                        var position2 = str2_tmp.length;
                                    }
                                }
                            }
                        }
                    }

                    //get position where to insert extracted value
                    var column_char = getCharFromNumber(header + 1);

                    //get value to extract
                    if (position2 > position1) {
                        var extractedValue = range.text[i][header].substring(position1, position2);
                        extract_count += 1;
                    }
                    else {
                        var extractedValue = "";
                        empty_count += 1;
                    }

                    var extract_tmp = [];
                    extract_tmp.push(extractedValue);
                    extracted_array.push(extract_tmp);

                }

                var column_char = getCharFromNumber(header + 1);
                var rangeaddress = column_char + ":" + column_char;
                var range_insert = ctx.workbook.worksheets.getActiveWorksheet().getRange(rangeaddress);
                range_insert.insert("Right");

                var insert_address = column_char + 1 + ":" + column_char + range.text.length;
                addExtractedValue(extracted_array, insert_address);

                var txt = document.createElement("p");
                txt.className = "ms-font-xs ms-embedded-dialog__content__text";
                txt.innerHTML = "PrepJet extracted " + extract_count + " values. " + empty_count + " data entries did not contain the specified delimiter or delimiter position."
                document.getElementById('resultText').appendChild(txt);

                document.getElementById('resultDialog').style.visibility = 'visible';

                //window.location = "extract_values.html";
            });


        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }


    function addExtractedValue(split_array, insert_address){

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('text');
            worksheet.load('name');

            return ctx.sync().then(function() {
                addContentNew(worksheet.name, insert_address, split_array);
            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

    }



})();