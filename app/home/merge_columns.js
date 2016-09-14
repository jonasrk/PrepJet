function backToOne() {
    $('#step1').show();
    $('#step2').hide();
    Office.context.document.settings.set('back_button_pressed', false);
    Office.context.document.settings.set('populate_new', true);
}


function redirectHome() {
    window.location = "mac_start.html";
}


(function () {
    // 'use strict';
    var count_drop = 0;

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            Office.context.document.settings.set('same_header_addcolumn', false);
            //save function to redirect to correct screen after intro
            Office.context.document.settings.set('last_clicked_function', "merge_columns.html");
            if (Office.context.document.settings.get('prepjet_loaded_before') == null) {
                Office.context.document.settings.set('backup_sheet_count', 1);
                Office.context.document.settings.set('prepjet_loaded_before', true);
                Office.context.document.settings.saveAsync();
                window.location = "intro.html";
            }


            Office.context.document.settings.set('back_button_pressed', false);

            app.initialize();


            $('#step2').hide();
            $('#step3').hide();
            $('#step4').hide();

            populateDropdowns();

            $('#bt_step2').click(step2ButtonClicked);
            $('#bt_step4').click(step4ButtonClicked);
            $('#bt_step3').click(step3Show); //todo: add function
            $('#back_step1').click(backToOne);
            $('#bt_apply').click(applyButtonClicked);
            $('#back_step2').click(step2ButtonClicked);
            $('#back_step3').click(step3Show); //todo: add function
            $('#buttonOk').click(highlightHeader);
            $('#homeButton').click(redirectHome);


            //show and hide error message for columns that have same header name
            document.getElementById("buttonClose").onclick = function () {
                document.getElementById('showEmbeddedDialog').style.visibility = 'hidden';
            }


            //show and hide help callouts
            document.getElementById("help_icon").onclick = function () {
                document.getElementById('helpCallout').style.visibility = 'visible';
            }
            document.getElementById("closeCallout").onclick = function () {
                document.getElementById('helpCallout').style.visibility = 'hidden';
            }
            document.getElementById("help_iconFirst").onclick = function () {
                document.getElementById('helpCalloutFirst').style.visibility = 'visible';
            }

            document.getElementById("closeCalloutFirst").onclick = function () {
                document.getElementById('helpCalloutFirst').style.visibility = 'hidden';
            }

            document.getElementById("refresh_icon").onclick = function () {
                window.location = "merge_columns.html";
            }

            //hide result message
            document.getElementById("resultClose").onclick = function () {
                document.getElementById('resultDialog').style.visibility = 'hidden';
                window.location = "merge_columns.html";
            }
            document.getElementById("resultOk").onclick = function () {
                document.getElementById('resultDialog').style.visibility = 'hidden';
                window.location = "merge_columns.html";
            }


        });
    };


    function checkCheckbox() {

        var selected_table2 = document.getElementById('table2_options').value;
        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getItem(selected_table2);
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange(true);
            var firstCell = range.getColumn(0);
            var firstCol = firstCell.getEntireColumn();
            var tmpRow = range.getRow(0);
            var firstRow = tmpRow.getEntireRow();


            range.load('text');
            firstRow.load('address');
            firstCol.load('address');

            return ctx.sync().then(function() {

                var tmp_offset = firstCol.address;
                var col_offset = tmp_offset.substring(tmp_offset.indexOf("!") + 1, tmp_offset.indexOf(":"));
                var tmp_row = firstRow.address;
                var row_offset = Number(tmp_row.substring(tmp_row.indexOf("!") + 1, tmp_row.indexOf(":")));
                var add_col = getNumberFromChar(col_offset);

                if (document.getElementById('checkbox_all').checked == true) {
                    for (var i = 0; i < range.text[0].length; i++) {
                        if (range.text[0][i] != "") {
                            document.getElementById(range.text[0][i]).checked = true;
                        }
                        else {
                            document.getElementById("Column " + getCharFromNumber(i + add_col)).checked = true;
                        }
                    }
                }
                else {
                    for (var i = 0; i < range.text[0].length; i++) {
                        if (range.text[0][i] != "") {
                            document.getElementById(range.text[0][i]).checked = false;
                        }
                        else {
                            document.getElementById("Column " + getCharFromNumber(i + add_col)).checked = false;
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
                                $("span.ms-Dropdown-title:empty").text(worksheet_names[0]);

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



    function highlightHeader() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange(true);
            var firstCell = range.getColumn(0);
            var firstCol = firstCell.getEntireColumn();
            var tmpRow = range.getRow(0);
            var firstRow = tmpRow.getEntireRow();

            range.load('text');
            firstRow.load('address');
            firstCol.load('address');
            worksheet.load('name');

            return ctx.sync().then(function() {

                var tmp_offset = firstCol.address;
                var col_offset = tmp_offset.substring(tmp_offset.indexOf("!") + 1, tmp_offset.indexOf(":"));
                var tmp_row = firstRow.address;
                var row_offset = Number(tmp_row.substring(tmp_row.indexOf("!") + 1, tmp_row.indexOf(":")));
                var add_col = getNumberFromChar(col_offset);

                for (var run = 0; run < range.text[0].length - 1; run++) {
                    for (var run2 = run + 1; run2 < range.text[0].length; run2++) {
                        if (range.text[0][run] == range.text[0][run2] && range.text[0][run] != "") {
                            document.getElementById('showEmbeddedDialog').style.visibility = 'hidden';
                            highlightContentNew(worksheet.name, getCharFromNumber(run + add_col) + row_offset, '#EA7F04', function () {});
                            highlightContentNew(worksheet.name, getCharFromNumber(run2 + add_col) + row_offset, '#EA7F04', function () {});
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



    function step2ButtonClicked() {

        $('#step1').hide();
        $('#step2').show();
        $('#step3').hide();
        $('#step4').hide();


        var selected_table2 = document.getElementById('table2_options').value; // TODO better reference by ID than name

        Excel.run(function (ctx) {

            var myBindings = Office.context.document.bindings;
            var worksheet = ctx.workbook.worksheets.getItem(selected_table2);
            var worksheetname = ctx.workbook.worksheets.getItem(selected_table2);
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange(true);
            var firstCell = range.getColumn(0);
            var firstCol = firstCell.getEntireColumn();
            var tmpRow = range.getRow(0);
            var firstRow = tmpRow.getEntireRow();

            firstRow.load('address');
            firstCol.load('address');
            range.load('text');
            worksheetname.load('name');

            return ctx.sync().then(function() {

                var tmp_offset = firstCol.address;
                var col_offset = tmp_offset.substring(tmp_offset.indexOf("!") + 1, tmp_offset.indexOf(":"));
                var tmp_row = firstRow.address;
                var row_offset = Number(tmp_row.substring(tmp_row.indexOf("!") + 1, tmp_row.indexOf(":")));
                var add_col = getNumberFromChar(col_offset);

                for (var run = 0; run < range.text[0].length - 1; run++) {
                    for (var run2 = run + 1; run2 < range.text[0].length; run2++) {
                        if (range.text[0][run] == range.text[0][run2] && range.text[0][run] != "") {
                            document.getElementById('showEmbeddedDialog').style.visibility = 'visible';
                            Office.context.document.settings.set('same_header_addcolumn', true);
                        }
                    }
                }

                //document.getElementById('checkbox_all').checked = false;
                while (document.getElementById('checkboxes_variables').firstChild) {
                    document.getElementById('checkboxes_variables').removeChild(document.getElementById('checkboxes_variables').firstChild);
                }

                for (var i = 0; i < range.text[0].length; i++) { // .text[0] is the first row of a range
                    if (range.text[0][i] != ""){
                        addNewCheckboxToContainer (range.text[0][i], "reference_column_checkbox" ,"checkboxes_variables");
                    }
                    else {
                        var colchar = getCharFromNumber(i + add_col);
                        addNewCheckboxToContainer ("Column " + colchar, "reference_column_checkbox" ,"checkboxes_variables");
                    }
                }
                $('#checkbox_all').click(checkCheckbox);


                //function to check whether header entries are changed
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
                    window.location = "merge_columns.html";
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
    }


    function step3Show() {

        $('#step1').hide();
        $('#step2').hide();
        $('#step3').show();
        $('#step4').hide();

        Excel.run(function (ctx) {

            return ctx.sync().then(function() {

            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }


    function step4ButtonClicked() {
        $('#step1').hide();
        $('#step2').hide();
        $('#step3').hide();
        $('#step4').show();
        $('#bt_remove').hide();


        var selected_table1 = document.getElementById('table1_options').value; // TODO better reference by ID than name
        var selected_table2 = document.getElementById('table2_options').value; // TODO better reference by ID than name


        function populateReferenceColumnDropdown (table1, table2, container_tmp) {

            if (Office.context.document.settings.get('back_button_pressed') == true) {
                var parentdiv = document.getElementById('dropdowns_step3');
                while (parentdiv.firstChild) {
                    parentdiv.removeChild(parentdiv.firstChild);
                }
                count_drop = 0;
            }

            Excel.run(function (ctx) {

                var worksheet_t1 = ctx.workbook.worksheets.getItem(table1);
                var range_all_t1 = worksheet_t1.getRange();
                var range_t1 = range_all_t1.getUsedRange(true);
                var firstCell1 = range_t1.getColumn(0);
                var firstCol1 = firstCell1.getEntireColumn();
                var tmpRow1 = range_t1.getRow(0);
                var firstRow1 = tmpRow1.getEntireRow();

                var worksheet_t2 = ctx.workbook.worksheets.getItem(table2);
                var range_all_t2 = worksheet_t2.getRange();
                var range_t2 = range_all_t2.getUsedRange(true);
                var firstCell2 = range_t2.getColumn(0);
                var firstCol2 = firstCell2.getEntireColumn();
                var tmpRow2 = range_t2.getRow(0);
                var firstRow2 = tmpRow2.getEntireRow();

                range_t1.load('address');
                range_t1.load('text');
                firstRow1.load('address');
                firstCol1.load('address');

                range_t2.load('address');
                range_t2.load('text');
                firstRow2.load('address');
                firstCol2.load('address');

                return ctx.sync().then(function() {

                    var tmp_offset1 = firstCol1.address;
                    var col_offset1 = tmp_offset1.substring(tmp_offset1.indexOf("!") + 1, tmp_offset1.indexOf(":"));
                    var tmp_row1 = firstRow1.address;
                    var row_offset1 = Number(tmp_row1.substring(tmp_row1.indexOf("!") + 1, tmp_row1.indexOf(":")));
                    var add_col1 = getNumberFromChar(col_offset1);

                    var tmp_offset2 = firstCol2.address;
                    var col_offset2 = tmp_offset2.substring(tmp_offset2.indexOf("!") + 1, tmp_offset2.indexOf(":"));
                    var tmp_row2 = firstRow2.address;
                    var row_offset2 = Number(tmp_row2.substring(tmp_row2.indexOf("!") + 1, tmp_row2.indexOf(":")));
                    var add_col2 = getNumberFromChar(col_offset2);

                    if (Office.context.document.settings.get('populate_new') == false) {
                        var count_tmp = count_drop + 1;
                    }
                    else {
                        var count_tmp = count_drop + 3;
                    }

                    var trow = document.createElement("tr");
                    trow.id = "lookuprow" + (count_drop + 1)
                    document.getElementById('matchCriteria').appendChild(trow);

                    for (var k = (count_drop + 1); k < count_tmp; k++) {
                        var container = container_tmp + k;
                        var div = document.createElement("div");
                        div.className = "ms-Dropdown reference_column_checkboxes_" + k;
                        div.id = "addedDropdown" + k;

                        var sel = document.createElement("select");
                        sel.id = container;
                        sel.className = "ms-Dropdown-select";

                        var lab = document.createElement('label');
                        lab.className = "ms-Label";
                        lab.setAttribute("for", "addedDropdown" + k);

                        var elemi = document.createElement("i");
                        elemi.className = "ms-Dropdown-caretDown ms-Icon ms-Icon--caretDown";

                        var tcol = document.createElement("td");
                        tcol.id = "lookuptable";
                        trow.appendChild(tcol);

                        //document.getElementById("dropdowns_step3").appendChild(div);
                        tcol.appendChild(div);
                        document.getElementById("addedDropdown" + k).appendChild(lab);
                        document.getElementById("addedDropdown" + k).appendChild(elemi);
                        document.getElementById("addedDropdown" + k).appendChild(sel);

                        if (k % 2 == 0) {
                            lab.innerHTML = table2;
                            for (var i = 0; i < range_t2.text[0].length; i++) {
                                var el = document.createElement("option");
                                if (range_t2.text[0][i] != "") {
                                    el.value = range_t2.text[0][i];
                                    el.textContent = range_t2.text[0][i];
                                }
                                else {
                                    el.value = "Column " + getCharFromNumber(i + add_col2);
                                    el.textContent = "Column " + getCharFromNumber(i + add_col2);
                                }
                                sel.appendChild(el);
                            }
                        }
                        else {
                            lab.innerHTML = table1;
                            for (var i = 0; i < range_t1.text[0].length; i++) {
                                var el = document.createElement("option");
                                if (range_t1.text[0][i] != "") {
                                    el.value = range_t1.text[0][i];
                                    el.textContent = range_t1.text[0][i];
                                }
                                else {
                                    el.value = "Column " + getCharFromNumber(i + add_col1);
                                    el.textContent = "Column " + getCharFromNumber(i + add_col1);
                                }

                                sel.appendChild(el);
                            }
                        }

                        document.getElementById("addedDropdown" + k).appendChild(lab);
                        $("." + container).Dropdown();
                        if (k % 2 == 0) {
                            $("span.ms-Dropdown-title:empty").text(range_t2.text[0][0]);
                        }
                        else {
                            $("span.ms-Dropdown-title:empty").text(range_t1.text[0][0]);
                        }
                        count_drop = count_drop + 1;
                    }
                    Office.context.document.settings.set('back_button_pressed', false);
                    Office.context.document.settings.set('populate_new', false);
                });

            }).catch(function(error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });

        }

        populateReferenceColumnDropdown(selected_table1, selected_table2, "reference_column_checkboxes_");

        $("#bt_more").unbind('click');
        $('#bt_more').click(addDropdown);
        $('#bt_remove').click(removeCriteria);


        function addDropdown(){
            $('#bt_remove').show();
            Office.context.document.settings.set('populate_new', true);
            populateReferenceColumnDropdown(selected_table1, selected_table2, "reference_column_checkboxes_");
        }


        function removeCriteria() {
            var loop_end = count_drop - 1;
            var parent = document.getElementById("matchCriteria");
            var child = document.getElementById("lookuprow" + loop_end);
            parent.removeChild(child);
            count_drop = count_drop - 2;
            if (count_drop < 3) {
                $('#bt_remove').hide();
            }
        }

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

            //ranges for source worksheet
            var worksheet = ctx.workbook.worksheets.getItem(selected_table2);
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange(true);
            var firstCell = range.getColumn(0);
            var firstCol = firstCell.getEntireColumn();
            var tmpRow = range.getRow(0);
            var firstRow = tmpRow.getEntireRow();

            firstRow.load('address');
            firstCol.load('address');
            range.load('address');
            range.load('text');

            //ranges for target working sheet
            var worksheet_adding_to = ctx.workbook.worksheets.getItem(selected_table1);
            var range_all_adding_to = worksheet_adding_to.getRange();
            var range_adding_to = range_all_adding_to.getUsedRange(true);
            var firstCellTarget = range_adding_to.getColumn(0);
            var firstColTarget = firstCellTarget.getEntireColumn();
            var tmpRowTarget = range_adding_to.getRow(0);
            var firstRowTarget = tmpRowTarget.getEntireRow();

            firstRowTarget.load('address');
            firstColTarget.load('address');
            range_adding_to.load('address');
            range_adding_to.load('text');
            worksheet_adding_to.load('name');

            Office.context.document.settings.set('populate_new', true);
            Office.context.document.settings.set('back_button_pressed', false);

            return ctx.sync().then(function() {

                var tmp_offsetTarget = firstColTarget.address;
                var col_offsetTarget = tmp_offsetTarget.substring(tmp_offsetTarget.indexOf("!") + 1, tmp_offsetTarget.indexOf(":"));
                var tmp_rowTarget = firstRowTarget.address;
                var row_offsetTarget = Number(tmp_rowTarget.substring(tmp_rowTarget.indexOf("!") + 1, tmp_rowTarget.indexOf(":")));
                var add_colTarget = getNumberFromChar(col_offsetTarget);

                var tmp_offsetSource = firstCol.address;
                var col_offsetSource = tmp_offsetSource.substring(tmp_offsetSource.indexOf("!") + 1, tmp_offsetSource.indexOf(":"));
                var tmp_rowSource = firstRow.address;
                var row_offsetSource = Number(tmp_rowSource.substring(tmp_rowSource.indexOf("!") + 1, tmp_rowSource.indexOf(":")));
                var add_colSource = getNumberFromChar(col_offsetSource);

                var startCell = col_offsetTarget + row_offsetTarget;

                backupForUndo(range_adding_to, startCell, add_colTarget, row_offsetTarget);

                var aggregation = document.getElementById('aggregation_options').value;

                var column1_ids = []; //new Array(identifier_length);
                var column2_ids = []; //new Array(identifier_length);

                //get vector with column indices of matcher for each table
                var pos_col1 = 0;
                var pos_col2 = 0;

                for (var runid1 = 0; runid1 < identifier1.length; runid1++) {
                    for (var runheader = 0; runheader < range_adding_to.text[0].length; runheader++){
                        if (identifier1[runid1] == range_adding_to.text[0][runheader] || identifier1[runid1] == "Column " + getCharFromNumber(runheader + add_colTarget)){
                            column1_ids[runid1] = runheader;
                        }
                    }
                }

                for (var runid2 = 0; runid2 < identifier2.length; runid2++) {
                    for (var runheader = 0; runheader < range.text[0].length; runheader++){
                        if (identifier2[runid2] == range.text[0][runheader] || identifier2[runid2] == "Column " + getCharFromNumber(runheader + add_colSource)){
                            column2_ids[runid2] = runheader;
                        }
                    }
                }

                var lookup_count = 0;
                var empty_count = 0;

                for (var k = 0; k < range.text[0].length; k++){

                    // iterate over checked checkboxes
                    var checked_checkboxes = getCheckedBoxes("reference_column_checkbox");

                    if (document.getElementById("case_sens").checked == true) {
                        var case_sens = 1;
                    }
                    else {
                        var case_sens = 0;
                    }

                    var source_char = getCharFromNumber(k + add_colSource);

                    for (var l = 0; l < checked_checkboxes.length; l++){ // TODO throws error if none are checked
                        if (checked_checkboxes[l].id == range.text[0][k] || checked_checkboxes[l].id == "Column " + getCharFromNumber(k)){
                            var lookup_array = [];
                            var column_char = getCharFromNumber(l + range_adding_to.text[0].length + add_colTarget);

                            // copy title
                            var headerText = ["=" + selected_table2 + "!" + source_char + row_offsetSource];
                            lookup_array.push(headerText);

                            // copy rest
                            for (var i = 1; i < range_adding_to.text.length; i++) {
                                var singleMatchCount = 0;
                                for (var j = 1; j < range.text.length; j++) {
                                    var check = 0;
                                    for (var runid = 0; runid < column1_ids.length; runid ++) {
                                        var col1 = column1_ids[runid];
                                        var col2 = column2_ids[runid];

                                        if (case_sens == 1) {
                                            if (range_adding_to.text[i][col1] == range.text[j][col2]) {
                                                check = check + 1;
                                            }
                                        }
                                        else {
                                            if (range_adding_to.text[i][col1].toLowerCase() == range.text[j][col2].toLowerCase()) {
                                                check = check + 1;
                                            }
                                        }

                                    }
                                    var check_match = 0;
                                    if (check == column1_ids.length) {
                                        var sheet_row = i + row_offsetTarget;
                                        var row_ref = row_offsetSource + j;
                                        //var textToAdd = ["=" + selected_table2 + "!" + source_char + row_ref];
                                        if (aggregation == "noagg") {
                                            var textToAdd = ["=" + selected_table2 + "!" + source_char + row_ref];
                                            lookup_array.push(textToAdd);
                                            lookup_count += 1;
                                            check_match = 1;
                                            break;
                                        }
                                        if (aggregation == "sum") {
                                            if (singleMatchCount == 0) {
                                                var textToAdd = ["=" + selected_table2 + "!" + source_char + row_ref];
                                            } else {
                                                textToAdd = [textToAdd + "+" + selected_table2 + "!" + source_char + row_ref];
                                            }
                                            check_match = 1;
                                        }
                                        singleMatchCount += 1;
                                    }
                                }
                                console.log(textToAdd);
                                console.log(check_match);
                                if (check_match == 0 && singleMatchCount == 0) {
                                    lookup_array.push([""]);
                                }
                                if (singleMatchCount != 0 && aggregation != "noagg") {
                                    lookup_array.push(textToAdd);
                                    lookup_count += 1;
                                }
                            }
                            var insert_address = column_char + 1 + ":" + column_char + range_adding_to.text.length;
                            addContentNew(worksheet_adding_to.name, insert_address, lookup_array, function(){});
                        }
                    }
                }


                if (document.getElementById('createBackup').checked == true) {
                    var sheet_count = Office.context.document.settings.get('backup_sheet_count') + 1;
                    Office.context.document.settings.set('backup_sheet_count', sheet_count);
                    Office.context.document.settings.saveAsync();
                    var newName = worksheet_adding_to.name + "(" + sheet_count + ")";
                    addBackupSheet(newName, startCell, add_colTarget, row_offsetTarget, function() {
                        empty_count = checked_checkboxes.length * range_adding_to.text.length - lookup_count - checked_checkboxes.length;
                        var txt = document.createElement("p");
                        txt.className = "ms-font-xs ms-embedded-dialog__content__text";
                        txt.innerHTML = "PrepJet found " + lookup_count + " matching data records. " + empty_count + " rows did not meet the specified match criteria."
                        document.getElementById('resultText').appendChild(txt);
                        document.getElementById('resultDialog').style.visibility = 'visible';
                    });
                }
                else {
                    empty_count = checked_checkboxes.length * range_adding_to.text.length - lookup_count - checked_checkboxes.length;
                    var txt = document.createElement("p");
                    txt.className = "ms-font-xs ms-embedded-dialog__content__text";
                    txt.innerHTML = "PrepJet found " + lookup_count + " matching data records. " + empty_count + " rows did not meet the specified match criteria."
                    document.getElementById('resultText').appendChild(txt);
                    document.getElementById('resultDialog').style.visibility = 'visible';
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