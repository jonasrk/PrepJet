function redirectHome() {
    window.location = "mac_start.html";
}

(function () {
    count_drop = 0;
    // 'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            Office.context.document.settings.set('same_header_outlier', false);
            Office.context.document.settings.set('last_clicked_function', "outlier.html");
            if (Office.context.document.settings.get('prepjet_loaded_before') == null) {
                Office.context.document.settings.set('backup_sheet_count', 1);
                Office.context.document.settings.set('prepjet_loaded_before', true);
                Office.context.document.settings.saveAsync();
                window.location = "intro.html";
            }

            app.initialize();
            populateColumnDropdown();

            $('#removeVar').hide();

            $('#bt_detect_outliers').click(detectOutlier);
            $('#homeButton').click(redirectHome);
            $('#buttonOk').click(highlightHeader);
            $('#addVar').click(addDropdown);
            $('#removeVar').click(removeCriteria);

            //refresh window
            document.getElementById("refresh_icon").onclick = function () {
                window.location = "outlier.html";
            }

            //Show and hide error message if column have same header name
            document.getElementById("buttonClose").onclick = function () {
                document.getElementById('showEmbeddedDialog').style.visibility = 'hidden';
            }

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

    function addDropdown(){
        $('#removeVar').show();
        count_drop += 1;
        populateDependendVariable();
    }

    function populateDependendVariable() {

            Excel.run(function (ctx) {

                var worksheet_t1 = ctx.workbook.worksheets.getActiveWorksheet();
                var range_all_t1 = worksheet_t1.getRange();
                var range_t1 = range_all_t1.getUsedRange();

                range_t1.load('address');
                range_t1.load('text');

                return ctx.sync().then(function() {

                        var div = document.createElement("div");
                        div.className = "ms-Dropdown reference_column_checkboxes" + count_drop;
                        div.id = "addedDropdown" + count_drop;

                        var sel = document.createElement("select");
                        sel.id = "dependendVariable" + count_drop;
                        sel.className = "ms-Dropdown-select";

                        var lab = document.createElement('label');
                        lab.className = "ms-Label";
                        lab.innerHTML = "Select column of dependend variable"
                        lab.setAttribute("for", "addedDropdown" + count_drop);

                        var elemi = document.createElement("i");
                        elemi.className = "ms-Dropdown-caretDown ms-Icon ms-Icon--caretDown";

                        document.getElementById('dependendVariables').appendChild(div);
                        document.getElementById("addedDropdown" + count_drop).appendChild(lab);
                        document.getElementById("addedDropdown" + count_drop).appendChild(elemi);
                        document.getElementById("addedDropdown" + count_drop).appendChild(sel);

                            for (var i = 0; i < range_t1.text[0].length; i++) {
                                var el = document.createElement("option");
                                if (range_t1.text[0][i] != "") {
                                    el.value = range_t1.text[0][i];
                                    el.textContent = range_t1.text[0][i];
                                }
                                else {
                                    el.value = "Column " + getCharFromNumber(i);
                                    el.textContent = "Column " + getCharFromNumber(i);
                                }

                                sel.appendChild(el);
                            }

                        document.getElementById("addedDropdown" + count_drop).appendChild(lab);
                        $(".reference_column_checkboxes" + count_drop).Dropdown();
                        $("span.ms-Dropdown-title:empty").text(range_t1.text[0][0]);
                });

            }).catch(function(error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });

    }

    function removeCriteria() {
            var loop_end = count_drop - 1;
            var parent = document.getElementById('dependendVariables');
            var child = document.getElementById('addedDropdown' + count_drop);
            parent.removeChild(child);
            count_drop = count_drop - 1;
            if (count_drop < 1) {
                $('#removeVar').hide();
            }
    }


    function populateColumnDropdown() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('address');
            range.load('text');
            return ctx.sync().then(function() {

                for (var run = 0; run < range.text[0].length - 1; run++) {
                    for (var run2 = run + 1; run2 < range.text[0].length; run2++) {
                        if (range.text[0][run] == range.text[0][run2] && range.text[0][run] != "") {
                            document.getElementById('showEmbeddedDialog').style.visibility = 'visible';
                            Office.context.document.settings.set('same_header_outlier', true);
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
                    document.getElementById("outlier_column_dropdown").appendChild(el);
                }

                $(".outlier_column_dropdown_container").Dropdown();
                $("span.ms-Dropdown-title:empty").text(range.text[0][0]);
            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

    }


    function detectOutlier() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('address');
            range.load('text');
            range.load('values');

            var selected_identifier = document.getElementById('outlier_column_dropdown').value;

            return ctx.sync().then(function() {

                var header = 0;
                for (var k = 0; k < range.text[0].length; k++){
                    if (selected_identifier == range.text[0][k] || selected_identifier == "Column " + getCharFromNumber(k)){
                        header = k;
                    }
                }

                if (count_drop > 0) {

                    var independent_identifier = document.getElementById('addedDropdown' + count_drop).value;
                    var headerIndep = 0;

                    for (var k = 0; k < range.text[0].length; k++){
                        if (independent_identifier == range.text[0][k] || independent_identifier == "Column " + getCharFromNumber(k)){
                            headerIndep = k;
                        }
                    }

                    var independent_array = [];
                    for (var i = 1; i < range.text.length; i++) {
                        var row_number = i + 1;
                        independent_array.push([range.values[i][headerIndep], range.values[i][header]]);
                    }

                    var data_set = [independent_array];

                    $.post( "https://localhost:8100/outlierIndep/", { data: data_set })
                    .done(function( borders ) {
                        // highlight dupes
                        console.log("Borders: " + borders + "\nStatus: " + status);

                        Excel.run(function (ctx) {

                            var dupe_worksheet = ctx.workbook.worksheets.getActiveWorksheet();
                            var dupe_range_all = dupe_worksheet.getRange();
                            var dupe_range = dupe_range_all.getUsedRange();

                            dupe_range.load('address');
                            dupe_range.load('text');

                            var selected_column = document.getElementById('outlier_column_dropdown').value;

                            return ctx.sync().then(function() {

                                var upper_border = borders['objects'][1];
                                var lower_border = borders['objects'][0];

                                console.log(upper_border);
                                console.log(lower_border);


                            });
                        });
                    });

                } else {

                    var data_array  = [];

                    for (var i = 1; i < range.text.length; i++) {
                        var row_number = i + 1;
                        data_array.push(range.values[i][header]);
                    }

                    $.post( "https://localhost:8100/outlier/", { data: data_array })
                    .done(function( borders ) {
                        // highlight dupes
                        console.log("Borders: " + borders + "\nStatus: " + status);

                        Excel.run(function (ctx) {

                            var dupe_worksheet = ctx.workbook.worksheets.getActiveWorksheet();
                            var dupe_range_all = dupe_worksheet.getRange();
                            var dupe_range = dupe_range_all.getUsedRange();

                            dupe_range.load('address');
                            dupe_range.load('text');

                            var selected_column = document.getElementById('outlier_column_dropdown').value;

                            return ctx.sync().then(function() {

                                var upper_border = borders['objects'][1];
                                var lower_border = borders['objects'][0];

                                var header = 0;
                                for (var k = 0; k < range.text[0].length; k++){
                                    if (selected_column == range.text[0][k] || selected_column == "Column " + getCharFromNumber(k)){
                                        header = k;
                                    }
                                }

                                var color = "#EA7F04";
                                for (var k = 1; k < dupe_range.text.length; k++) {
                                    if (dupe_range.text[k][header] < lower_border || dupe_range.text[k][header] > upper_border) {
                                        var insert_address = getCharFromNumber(header) + (k + 1);
                                        console.log(insert_address);
                                        highlightCellInWorksheet(dupe_worksheet, insert_address, color);
                                    }
                                }

                            });
                        });
                    });
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