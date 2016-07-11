function backToOne() {
    $('#step1').show();
    $('#step2').hide();
}

(function () {
    'use strict';
    var count_drop = 2;
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            Office.context.document.settings.set('last_clicked_function', "merge_columns.html");
            if (Office.context.document.settings.get('prepjet_loaded_before') == null) {
                Office.context.document.settings.set('prepjet_loaded_before', true);
                Office.context.document.settings.saveAsync();
                window.location = "intro.html";
            }


            app.initialize();

            $('#step2').hide();
            $('#step3').hide();

            populateDropdowns();

            $('#bt_step2').click(step2ButtonClicked);
            $('#bt_step3').click(step3ButtonClicked);
            $('#back_step1').click(backToOne);
            $('#bt_apply').click(applyButtonClicked);
            $('#back_step2').click(step2ButtonClicked);

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

                while (document.getElementById('checkboxes_variables').firstChild) {
                    document.getElementById('checkboxes_variables').removeChild(document.getElementById('checkboxes_variables').firstChild);
                }

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
        $('#bt_remove').hide();


        var selected_table1 = document.getElementById('table1_options').value; // TODO better reference by ID than name
        var selected_table2 = document.getElementById('table2_options').value; // TODO better reference by ID than name

        function populateReferenceColumnDropdown (table, container) {

        //remove potentially existing dropdown options
            var child_target = document.getElementById(container).firstChild;
            while (child_target != null) {
                document.getElementById(container).removeChild(child_target);
            }

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

        function addDropdown(){
            $('#bt_remove').show();
            var loop_end = count_drop + 1;
            for (var j = loop_end; j < (loop_end + 2); j++) {
                var div = document.createElement("div");
                div.className = "ms-Dropdown reference_column_checkboxes_" + j;

                var label = document.createElement("label");
                label.className = "ms-label";
                label.textContent = "Select reference column in table";

                var i = document.createElement("i");
                i.className = "ms-Dropdown-caretDown ms-Icon ms-Icon--caretDown";

                var select = document.createElement("select");
                select.className = "ms-Dropdown-select";
                select.id = "reference_column_checkboxes_" + j;

                div.appendChild(label);
                div.appendChild(i);
                div.appendChild(select);

                if (j % 2 == 0) {
                    var tmp_table = selected_table2;
                }
                else {
                    var tmp_table = selected_table1;
                }

                document.getElementById("dropdowns_step3").appendChild(div);
                populateReferenceColumnDropdown(tmp_table, "reference_column_checkboxes_" + j);
                count_drop = count_drop + 1;
            }
        }

        function removeCriteria() {
            var loop_end = count_drop - 1;
            for (var run = loop_end; run < (loop_end + 2); run++) {
                //var tmp = document.getElementById("reference_column_checkboxes_" + run);
                //tmp.style.display = 'none';
                $('#reference_column_checkboxes_' + run).hide();
            }
            count_drop = count_drop - 2;
        }

        populateReferenceColumnDropdown(selected_table1, "reference_column_checkboxes_1");
        populateReferenceColumnDropdown(selected_table2, "reference_column_checkboxes_2");

        $('#bt_more').click(addDropdown);
        $('#bt_remove').click(removeCriteria);

    }


    function applyButtonClicked() {
        $('#step1').hide();
        $('#step2').hide();
        $('#step3').hide();

        // find columns to match
        var identifier_length = count_drop / 2;
        var identifier1 = new Array(count_drop / 2);
        var identifier2 = new Array(count_drop / 2);

        var ident1_pos = 0;
        var ident2_pos = 0;

        for (var run = 0; run < count_drop; run++) {
            var countid = run + 1;
            if (countid % 2 != 0) {
                identifier1[ident1_pos] = document.getElementById("reference_column_checkboxes_" + countid).value; // TODO better reference by ID than name
                ident1_pos = ident1_pos + 1;
            }
            else {
                identifier2[ident2_pos] = document.getElementById("reference_column_checkboxes_" + countid).value; // TODO better reference by ID than name
                ident2_pos = ident2_pos + 1;
            }
        }

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

                var column1_ids = []; //new Array(identifier_length);
                var column2_ids = []; //new Array(identifier_length);

                //get vector with column indices of matcher for each table
                var pos_col1 = 0;
                var pos_col2 = 0;

                for (var runid1 = 0; runid1 < identifier1.length; runid1++) {
                    for (var runheader = 0; runheader < range_adding_to.text[0].length; runheader++){
                        if (identifier1[runid1] == range_adding_to.text[0][runheader]){
                            column1_ids[runid1] = runheader;
                        }
                    }
                }

                for (var runid2 = 0; runid2 < identifier2.length; runid2++) {
                    for (var runheader = 0; runheader < range.text[0].length; runheader++){
                        if (identifier2[runid2] == range.text[0][runheader]){
                            column2_ids[runid2] = runheader;
                        }
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
                            for (var i = 1; i < range_adding_to.text.length; i++) {
                                for (var j = 1; j < range.text.length; j++) {
                                    var check = 0;
                                    for (var runid = 0; runid < column1_ids.length; runid ++) {
                                        var col1 = column1_ids[runid];
                                        var col2 = column2_ids[runid];

                                        if (range_adding_to.text[i][col1] == range.text[j][col2]) {
                                            check = check + 1;
                                        }
                                    }
                                    if (check == column1_ids.length) {
                                        var sheet_row = i + 1;
                                        addContentToWorksheet(worksheet_adding_to, column_char + sheet_row, range.text[j][k])
                                        break; //todo correct position to ensure stops after one match found
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