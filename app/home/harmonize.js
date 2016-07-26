(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            Office.context.document.settings.set('last_clicked_function', "trim_spaces.html");
            if (Office.context.document.settings.get('prepjet_loaded_before') == null) {
                Office.context.document.settings.set('prepjet_loaded_before', true);
                Office.context.document.settings.saveAsync();
                window.location = "intro.html";
            }

            app.initialize();
            fillColumn();

            $(".harmonize_drop").Dropdown();

            $('#harmonize').click(harmonize);
            $('#checkbox_all').click(checkCheckbox);

        });
    };


    function checkCheckbox() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('text');

            return ctx.sync().then(function() {
                if (document.getElementById('checkbox_all').checked == true) {
                    for (var i = 0; i < range.text[0].length; i++) {
                        if (range.text[0][i] != "") {
                            document.getElementById(range.text[0][i]).checked = true;
                        }
                        else {
                            document.getElementById("Column " + getCharFromNumber(i)).checked = true;
                        }
                    }
                }
                else {
                    for (var i = 0; i < range.text[0].length; i++) {
                        if (range.text[0][i] != "") {
                            document.getElementById(range.text[0][i]).checked = false;
                        }
                        else {
                            document.getElementById("Column " + getCharFromNumber(i)).checked = false;
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
                for (var i = 0; i < range.text[0].length; i++) {
                    if (range.text[0][i] != ""){
                        addNewCheckboxToContainer (range.text[0][i], "column_checkbox" ,"checkboxes_columns");
                    }
                    else {
                        var colchar = getCharFromNumber(i);
                        addNewCheckboxToContainer ("Column " + colchar, "column_checkbox" ,"checkboxes_columns");
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


    function harmonize() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            //get used range in active Sheet
            range.load('text');

            var harmo = document.getElementById('harmonize_options').value;

            return ctx.sync().then(function() {

                backupForUndo(range);

                var header = 0;
                var act_worksheet = ctx.workbook.worksheets.getActiveWorksheet();
                var checked_checkboxes = getCheckedBoxes("column_checkbox");

                for (var run = 0; run < checked_checkboxes.length; run++) {
                    for (var k = 0; k < range.text[0].length; k++) {
                        if (checked_checkboxes[run].id == range.text[0][k] || checked_checkboxes[run].id == "Column " + getCharFromNumber(k)){
                            header = k;
                            break;
                        }
                    }

                    for (var k = 0; k < range.text.length; k++) {
                        if (harmo == "allupper") {
                            var harm_string = range.text[k][header].toUpperCase();
                        }
                        if (harmo == "alllower") {
                            var harm_string = range.text[k][header].toLowerCase();
                        }
                        if (harmo == "firstupper") { //todo when leading space first real letter not transformed
                            var tmp = range.text[k][header].toLowerCase().split(" ");
                            var tmp_upper = [];
                            for (var runtmp = 0; runtmp < tmp.length; runtmp++) {
                                tmp_upper.push(tmp[runtmp].charAt(0).toUpperCase() + tmp[runtmp].slice(1));
                            }
                            var harm_string = tmp_upper[0];
                            for (var runtmp = 1; runtmp < tmp_upper.length; runtmp++) {
                                harm_string = harm_string.concat(" ", tmp_upper[runtmp]);
                            }
                        }
                        if (harmo == "oneupper") {
                            var tmp = range.text[k][header].split(" ");
                            var tmp_upper = [];
                            tmp_upper.push(tmp[0].charAt(0).toUpperCase() + tmp[0].slice(1).toLowerCase());
                            for (var runtmp = 1; runtmp < tmp.length; runtmp++) {
                                tmp_upper.push(tmp[runtmp].charAt(0) + tmp[runtmp].slice(1).toLowerCase());
                            }

                            var harm_string = tmp_upper[0];
                            for (var runtmp = 1; runtmp < tmp_upper.length; runtmp++) {
                                harm_string = harm_string.concat(" ", tmp_upper[runtmp]);
                            }
                        }

                        var column_char = getCharFromNumber(header);
                        var sheet_row = k + 1;
                        addContentToWorksheet(act_worksheet, column_char + sheet_row, harm_string);

                    }
                }
                window.location = "harmonize.html";
            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }


})();