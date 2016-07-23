// The initialize function must be run each time a new page is loaded
(function () {
    Office.initialize = function (reason) {
        //If you need to initialize something you can do so here.
    };

})();

//Notice function needs to be in global namespace
function undo() { // TODO only does text, not formulas and formatting
    var backup_range_text = Office.context.document.settings.get('sheet_backup');

    Excel.run(function (ctx) {

        var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
        var range_all = worksheet.getRange();
        var range = range_all.getUsedRange();

        //get used range in active Sheet
        range.load('text');

        return ctx.sync().then(function() {

            var act_worksheet = ctx.workbook.worksheets.getActiveWorksheet();

            for (var i = 0; i < range.text.length; i++) {
                for (var j = 0; j < range.text[0].length; j++) {
                    var sheet_row = i + 1;
                    var column_char = getCharFromNumber(j);
                    addContentToWorksheet(act_worksheet, column_char + sheet_row, "");
                }
            }


            for (var i = 0; i < backup_range_text.length; i++) {
                for (var j = 0; j < backup_range_text[0].length; j++) {
                    var sheet_row = i + 1;
                    var column_char = getCharFromNumber(j);
                    addContentToWorksheet(act_worksheet, column_char + sheet_row, backup_range_text[i][j]);
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