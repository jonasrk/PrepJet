(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            Office.context.document.settings.set('last_clicked_function', "inconsistency.html");
            if (Office.context.document.settings.get('prepjet_loaded_before') == null) {
                Office.context.document.settings.set('prepjet_loaded_before', true);
                Office.context.document.settings.saveAsync();
                window.location = "intro.html";
            }

            app.initialize();
            fillColumn();


            $('#inconsistency').click(inconsistencies);
            $('#checkbox_all').click(checkCheckbox);

        });
    };


    function checkCheckbox() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('address');
            range.load('text');

            return ctx.sync().then(function() {
                if (document.getElementById('checkbox_all').checked == true) {
                    for (var i = 0; i < range.text[0].length; i++) {
                        if (range.text[0][i] != "") {
                            document.getElementById(range.text[0][i]).checked = true;
                        }
                        else {
                            document.getElementById("Column " + getCharFromNumber(i + 1)).checked = true;
                        }
                    }
                }
                else {
                    for (var i = 0; i < range.text[0].length; i++) {
                        if (range.text[0][i] != "") {
                            document.getElementById(range.text[0][i]).checked = false;
                        }
                        else {
                            document.getElementById("Column " + getCharFromNumber(i + 1)).checked = false;
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
                        var colchar = getCharFromNumber(i + 1);
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


    function inconsistencies() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('text');
            range.load('valueTypes'); //does not know date type
            range.load('values');

            return ctx.sync().then(function() {

                var header = 0;
                var checked_checkboxes = getCheckedBoxes("column_checkbox");
                var val_type = [];
                var check = [];
                var types = [];

                for (var run = 0;run < checked_checkboxes.length; run++) {
                    check[run] = 0;
                    var type_maximum = 0;
                    for (var k = 0; k < range.text[0].length; k++) {
                        if (checked_checkboxes[run].id == range.text[0][k] || checked_checkboxes[run].id == "Column " + getCharFromNumber(k+1)){
                            header = k;
                            break;
                        }
                    }

                    var tmp_type = [];
                    var tmp_address = "";
                    var rangeType = [];
                    var tmpUniqueTypes = [];
                    for (var i = 1; i < range.text.length; i++) {
                        var rangeType = [];
                        tmp_address = getCharFromNumber(header + 1) + (i + 1);
                        rangeType.push(range.valueTypes[i][header]);
                        if (i == 1) {
                            tmpUniqueTypes.push(range.valueTypes[i][header]);
                        }
                        rangeType.push(tmp_address);
                        tmp_type.push(rangeType);
                        if (i > 1 && (tmp_type[i - 1][0] != tmp_type[i - 2][0])) {
                            check[run] = 1;
                            var test_unique = 0
                            for (var k = 0; k < tmpUniqueTypes.length; k++) {
                                if (tmpUniqueTypes[k] != tmp_type[i - 1][0]) {
                                    test_unique = test_unique + 1;
                                }
                            }
                            if (test_unique >= tmpUniqueTypes.length) {
                                tmpUniqueTypes.push(tmp_type[i - 1][0]);
                            }
                        }
                    }
                    val_type.push(tmp_type);

                    if (check[run] == 1) {
                        var tmp2 = [];
                        for (var j = 0; j < tmpUniqueTypes.length; j++) {
                            var type_counter = 0;
                            var tmp1 = [];
                            for (var i = 0; i < tmp_type.length; i++) {
                                if (tmp_type[i][0] == tmpUniqueTypes[j]) {
                                    type_counter = type_counter + 1;
                                }
                            }

                            if (type_maximum < type_counter) {
                                type_maximum = type_counter;
                            }
                            //todo when 2 data types occure with them frequency none is highlighted
                            tmp1.push(tmpUniqueTypes[j]);
                            tmp1.push(type_counter);
                            tmp2.push(tmp1);
                        }

                        var color = "red";
                        for (var i = 0; i < tmp2.length; i++) {
                            if (tmp2[i][1] < type_maximum) {
                                for (var k = 0; k < tmp_type.length; k++) {
                                    if (tmp2[i][0] == tmp_type[k][0]) {
                                        highlightCellInWorksheet(worksheet, tmp_type[k][1], color);
                                    }
                                }
                            }
                        }

                        types.push(tmp2);
                    }

                }
                window.location = "inconsistency.html";
            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }


})();