(function () {
    // 'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            Office.context.document.settings.set('last_clicked_function', "undo_help.html");
            if (Office.context.document.settings.get('prepjet_loaded_before') == null) {
                Office.context.document.settings.set('prepjet_loaded_before', true);
                Office.context.document.settings.saveAsync();
                window.location = "intro.html";
            }

            app.initialize();

        });
    };


    function trimSpace() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            //get used range in active Sheet
            range.load('text');

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

                    for (var i = 1; i < range.text.length; i++) {
                        var trim_string = range.text[i][header].trim();
                        var column_char = getCharFromNumber(header);
                        var sheet_row = i + 1;
                        addContentToWorksheet(act_worksheet, column_char + sheet_row, trim_string);
                    }
                }
                window.location = "undo_help.html";
            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

})();