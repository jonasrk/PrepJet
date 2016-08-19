function redirectHome() {
    window.location = "mac_start.html";
}

(function () {
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

            $('#bt_detect_outliers').click(detectOutlier);
            $('#homeButton').click(redirectHome);

            document.getElementById("refresh_icon").onclick = function () {
                window.location = "harmonize.html";
            }

        });
    };


    function populateColumnDropdown() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('address');
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
                    document.getElementById("outlier_column_dropdown").appendChild(el);

                }

                $(".outlier_column_dropdown_container").Dropdown();
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

                var data_array  = [];

                for (var i = 1; i < range.text.length; i++) {
                    var row_number = i + 1;
                    data_array.push(range.values[i][header]);
                }


                // call to API

                $.post( "https://localhost:8100/outlier/", { data: data_array })
                    .done(function( borders ) {
                        // highlight dupes
                        console.log("Borders: " + borders + "\nStatus: " + status);
                        console.log(borders['objects'][0]);

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


            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

    }


})();