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

function displayAdvancedCount() {
    if (document.getElementById('advanced_settings').checked == true) {
        $('#del_count_start').show();
        $('.del_count_dropdown_s').show();
        $('#del_count_end').show();
        $('.del_count_dropdown_e').show();
    }
    else {
        $('#del_count_start').hide();
        $('.del_count_dropdown_s').hide();
        $('#del_count_end').hide();
        $('.del_count_dropdown_e').hide();
    }
}


function getColumn() {

    Excel.run(function (ctx) {

        var selectedRange = ctx.workbook.getSelectedRange();
        selectedRange.load('address');

        return ctx.sync().then(function() {
            document.getElementById('target_column_input').value = selectedRange.address;
        });

    }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
    });
}





(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

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
            $('.del_count_dropdown_s').Dropdown().hide();
            $('#del_count_end').hide();
            $('.del_count_dropdown_e').Dropdown().hide();

            $(".dropdown_table").Dropdown();
            $(".ms-TextField").TextField();

            $('#extract_Value').click(extractValue);


            Office.context.document.addHandlerAsync("documentSelectionChanged", myHandler, function(result){}
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
            }

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
            var target_tmp = document.getElementById('target_column_input').value
            if (target_tmp.indexOf(":") != -1) {
                var target_column = target_tmp.substring(target_tmp.indexOf("!") + 1, target_tmp.indexOf(":"));
            }
            else { //todo not correct to extract until +2 - better solution if only one column is selected
                var target_column = target_tmp.substring(target_tmp.indexOf("!") + 1, target_tmp.indexOf("!") + 2);
            }

            //if advanced settings are selected, get values for delimiter count
            if (document.getElementById('advanced_settings').checked == true) {
                var count_delimiter_start = Number(document.getElementById('delimiter_count_start').value);
                var count_direction_start = document.getElementById('del_count_drop_start').value;
                var count_delimiter_end = Number(document.getElementById('delimiter_count_end').value);
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
                var header = 0;
                //get column in header from which to extract value
                for (var k = 0; k < range.text[0].length; k++){
                    if (selected_identifier == range.text[0][k] || selected_identifier == "Column " + getCharFromNumber(k + 1)){
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
                    var sheet_row = i + 1;
                    if (target_column != "") {
                        var column_char = target_column
                    }
                    else {
                        var column_char = getCharFromNumber(header + 2);
                    }

                    //get value to extract
                    if (position2 > position1) {
                        var extractedValue = range.text[i][header].substring(position1, position2);
                    }
                    else {
                        var extractedValue = "";
                    }


                    //set position to insert extracted value
                    var rangeaddress = column_char + sheet_row;
                    var range_insert = ctx.workbook.worksheets.getActiveWorksheet().getRange(rangeaddress);
                    range_insert.insert("Right");
                    addContentToWorksheet(act_worksheet, column_char + sheet_row, extractedValue);

                }

                window.location = "extract_values.html";
            });


        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }



})();