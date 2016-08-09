// The initialize function must be run each time a new page is loaded
(function () {
    Office.initialize = function (reason) {
        //If you need to initialize something you can do so here.
    };

})();

//Notice function needs to be in global namespace
function undo() { // TODO only does text, not formulas and formatting
    Excel.run(function (ctx) {
        var values = Office.context.document.settings.get('sheet_backup');
        var end_address = getCharFromNumber(values[0].length - 1) + (values.length).toString();
        var rangeAddress = "A1:" + end_address;
        var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
        var range = worksheet.getRange(rangeAddress);

        var range_all = worksheet.getRange();
        var used_range = range_all.getUsedRange();

        used_range.clear();
        range.values = values;
        range.load('text');
        return ctx.sync().then(function() {
            // console.log(range.text);
        });
    }).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });

}
