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


    function populateDropdowns() {

        var worksheet_names = [];

        Excel.run(function (ctx) {
            var worksheets = ctx.workbook.worksheets;
            worksheets.load('items');
            return ctx.sync().then(function () {
                for (var i = 0; i < worksheets.items.length; i++) {
                    worksheets.items[i].load('name');
                    // worksheets.items[i].load('index'); TODO use index for something or do not load it
                    ctx.sync().then(function (i) {

                        var this_i = i;

                        return function () {
                            worksheet_names.push(worksheets.items[this_i].name);

                            if (worksheet_names.length == worksheets.items.length) {

                                for (var i = 0; i < worksheet_names.length; i++) { // TODO unnecessary loop
                                    var opt = worksheet_names[i];
                                    var el = document.createElement("option");
                                    el.textContent = opt;
                                    el.value = opt;
                                    document.getElementById("table1_options").appendChild(el);
                                    var el = document.createElement("option"); // TODO DRY
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


    function step2ButtonClicked() {

        $('#step1').hide();
        $('#step2').show();
        $('#step3').hide();

        var selected_table2 = document.getElementById('table2_options').value; // TODO better reference by ID than name

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getItem(selected_table2);
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('address');
            range.load('text');
            return ctx.sync().then(function() {
                for (var i = 0; i < range.text[0].length; i++) { // .text[0] is the first row of a range

                    addNewCheckboxToContainer (range.text[0][i], "reference_column_checkbox" ,"checkboxes_variables");
                }
            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function step3ButtonClicked() {
        $('#step1').hide();
        $('#step2').hide();
        $('#step3').show();

        var selected_table1 = document.getElementById('table1_options').value; // TODO better reference by ID than name
        var selected_table2 = document.getElementById('table2_options').value; // TODO better reference by ID than name

        function populateReferenceColumnDropdown (table, container) {

            Excel.run(function (ctx) {

                var worksheet = ctx.workbook.worksheets.getItem(table);
                var range_all = worksheet.getRange();
                var range = range_all.getUsedRange();

                range.load('address');
                range.load('text');
                return ctx.sync().then(function() {
                    for (var i = 0; i < range.text[0].length; i++) {

                        var el = document.createElement("option");
                        el.value = range.text[0][i];
                        el.textContent = range.text[0][i];
                        document.getElementById(container).appendChild(el);

                    }

                    $("." + container).Dropdown();
                });

            }).catch(function(error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });

        }

        populateReferenceColumnDropdown(selected_table1, "reference_column_ckeckboxes_1");
        populateReferenceColumnDropdown(selected_table2, "reference_column_ckeckboxes_2");
    }


    function applyButtonClicked() {
        $('#step1').hide();
        $('#step2').hide();
        $('#step3').hide();

        // find columns to match
        var selected_identifier1 = document.getElementById('reference_column_ckeckboxes_1').value; // TODO better reference by ID than name
        var selected_identifier2 = document.getElementById('reference_column_ckeckboxes_1').value; // TODO better reference by ID than name

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

                // initialize ids
                var sheet1_id = 0;
                var sheet2_id = 0;

                // iterate over columns

                for (var k = 0; k < range.text[0].length; k++){
                    if (selected_identifier1 == range.text[0][k]){
                        sheet1_id = k;
                    }
                }

                for (var k = 0; k < range_adding_to.text[0].length; k++){
                    if (selected_identifier2 == range_adding_to.text[0][k]){
                        sheet2_id = k;
                    }
                }

                for (var k = 0; k < range.text[0].length; k++){

                    // iterate over checked checkboxes

                    var checked_checkboxes = getCheckedBoxes("reference_column_checkbox");

                    for (var l = 0; l < checked_checkboxes.length; l++){ // TODO throws error if none are checked

                        if (checked_checkboxes[l].id == range.text[0][k]){

                            var column_char = getCharFromNumber(1 + l + range_adding_to.text[0].length);

                            // copy title
                            addContentToWorksheet(worksheet_adding_to, column_char + "1", range.text[0][k]);

                            // copy rest
                            for (var i = 1; i < range.text.length; i++) {
                                for (var j = 1; j < range_adding_to.text.length; j++) {
                                    if (range_adding_to.text[j][sheet2_id] == range.text[i][sheet1_id]) {
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

})();