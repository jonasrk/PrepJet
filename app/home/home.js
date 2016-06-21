(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {
            app.initialize();

            $('#step2').hide();
            $('#step3').hide();

            populateDropdowns();

            $('#bt_step2').click(step2ButtonClicked);
            $('#bt_step3').click(step3ButtonClicked);
            $('#bt_apply').click(applyButtonClicked);

        });
    };

    function step2ButtonClicked() {

        $('#step1').hide();
        $('#step2').show();
        $('#step3').hide();

        var selected_table2 = document.getElementById('table2_options').value; // TODO better reference by ID than name

        Excel.run(function (ctx) {
            var worksheet = ctx.workbook.worksheets.getItem(selected_table2);

            var rangeAddress = "A:Z"; // TODO Z is not the maximum
            var range_all = worksheet.getRange(rangeAddress);
            var range = range_all.getUsedRange();

            range.load('address');
            range.load('text');
            return ctx.sync().then(function() {
                for (var i = 0; i < range.text[0].length; i++) {

                    var el =  document.createElement("input");
                    el.name = "column_name_checkboxes";
                    el.id = range.text[0][i];
                    el.setAttribute("type", "checkbox");

                    var label = document.createElement("label");
                    label.textContent = range.text[0][i];
                    label.appendChild(el);

                    document.getElementById("checkboxes_variables").appendChild(label).appendChild(document.createElement("br"));
                }
            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    // Pass the checkbox name to the function
    function getCheckedBoxes(chkboxName) {
        var checkboxes = document.getElementsByName(chkboxName);
        var checkboxesChecked = [];
        // loop over them all
        for (var i=0; i<checkboxes.length; i++) {
            // And stick the checked ones onto an array...
            if (checkboxes[i].checked) {
                checkboxesChecked.push(checkboxes[i]);
            }
        }
        // Return the array if it is non-empty, or null
        return checkboxesChecked.length > 0 ? checkboxesChecked : null;
    }

    function step3ButtonClicked() {
        $('#step1').hide();
        $('#step2').hide();
        $('#step3').show();

        var selected_table1 = document.getElementById('table1_options').value; // TODO better reference by ID than name
        var selected_table2 = document.getElementById('table2_options').value; // TODO better reference by ID than name

        Excel.run(function (ctx) {
            var worksheet = ctx.workbook.worksheets.getItem(selected_table1);

            var rangeAddress = "A:Z"; // TODO Z is not the maximum
            var range_all = worksheet.getRange(rangeAddress);
            var range = range_all.getUsedRange();

            range.load('address');
            range.load('text');
            return ctx.sync().then(function() {
                for (var i = 0; i < range.text[0].length; i++) {

                    var el = document.createElement("option");
                    el.value = range.text[0][i];
                    el.textContent = range.text[0][i];
                    document.getElementById("reference_column_ckeckboxes_1").appendChild(el);

                }

                $(".reference_column_ckeckboxes_1").Dropdown();
            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

        Excel.run(function (ctx) {
            var worksheet = ctx.workbook.worksheets.getItem(selected_table2);

            var rangeAddress = "A:Z"; // TODO Z is not the maximum
            var range_all = worksheet.getRange(rangeAddress);
            var range = range_all.getUsedRange();

            range.load('address');
            range.load('text');
            return ctx.sync().then(function() {
                for (var i = 0; i < range.text[0].length; i++) {

                    var el = document.createElement("option");
                    el.value = range.text[0][i];
                    el.textContent = range.text[0][i];
                    document.getElementById("reference_column_ckeckboxes_2").appendChild(el);

                }

                $(".reference_column_ckeckboxes_2").Dropdown();
            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function applyButtonClicked() {
        $('#step1').show();
        $('#step2').hide();
        $('#step3').hide();

        var selected_table1 = document.getElementById('table1_options').value; // TODO better reference by ID than name
        var selected_table2 = document.getElementById('table2_options').value; // TODO better reference by ID than name

        Excel.run(function (ctx) {
            var worksheet = ctx.workbook.worksheets.getItem(selected_table2);

            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('address');
            range.load('text');

            var worksheet_adding_to = ctx.workbook.worksheets.getItem(selected_table1);

            var range_all_adding_to = worksheet_adding_to.getRange();
            var range_adding_to = range_all_adding_to.getUsedRange();

            range_adding_to.load('address');
            range_adding_to.load('text');

            return ctx.sync().then(function() {

                // iterate over columns

                for (var k = 0; k < range.text[0].length; k++){

                    // iterate over checked checkboxes

                    var checked_checkboxes = getCheckedBoxes("column_name_checkboxes");

                    for (var l = 0; l < checked_checkboxes.length; l++){

                        if (checked_checkboxes[l].id == range.text[0][k]){

                            // copy title

                            var column_char ='J';

                            if (l == 1) {
                                column_char ='K';
                            }

                            console.log("Match! Column Char: " + column_char);

                            addContentToWorksheet(worksheet_adding_to, column_char + "1", range.text[0][k]);

                            // copy rest

                            for (var i = 1; i < range.text.length; i++) {// TODO do not hardcode column

                                for (var j = 1; j < range_adding_to.text.length; j++) {

                                    // TODO do not hardcode column

                                    if (range_adding_to.text[j][8] == range.text[i][1]) {
                                        var sheet_row = j + 1;
                                        addContentToWorksheet(worksheet_adding_to, column_char + sheet_row, range.text[i][k])

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

    // Helper function to add and format content in the workbook
    function addContentToWorksheet(sheetObject, rangeAddress, displayText) {
        var range = sheetObject.getRange(rangeAddress);
        range.values = displayText;
        range.merge();
    }

    function populateDropdowns() {

        var allworksheets = [];

        Excel.run(function (ctx) {
            var worksheets = ctx.workbook.worksheets;
            worksheets.load('items');
            return ctx.sync().then(function () {
                for (var i = 0; i < worksheets.items.length; i++) {
                    worksheets.items[i].load('name');
                    worksheets.items[i].load('index');
                    ctx.sync().then(function (i) {

                        var this_i = i;

                        return function () {
                            allworksheets.push(worksheets.items[this_i].name);

                            if (this_i == worksheets.items.length - 1) {

                                for (var i = 0; i < allworksheets.length; i++) {
                                    var opt = allworksheets[i];
                                    var el = document.createElement("option");
                                    el.textContent = opt;
                                    el.value = opt;
                                    document.getElementById("table1_options").appendChild(el);
                                    var el = document.createElement("option");
                                    el.textContent = opt;
                                    el.value = opt;
                                    document.getElementById("table2_options").appendChild(el);
                                }

                                $(".dropdown_table").Dropdown();

                            }
                        }

                    }(i));
                }

            });

        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });


    }

})();
