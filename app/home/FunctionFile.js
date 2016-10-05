function redirectHome() {
    window.location = "mac_start.html";
}

// The initialize function must be run each time a new page is loaded
(function () {
    Office.initialize = function (reason) {
        undo();

        jQuery('#homeButton').click(redirectHome);

        //If you need to initialize something you can do so here.
    };

})();

//Notice function needs to be in global namespace
function undo() { // TODO only does text, not formulas and formatting
    Excel.run(function (ctx) {
        var values = Office.context.document.settings.get('sheet_backup');
        var startCell = Office.context.document.settings.get('startCell');
        var add_col = Office.context.document.settings.get('addCol');
        var row_offset  = Office.context.document.settings.get('rowOffset');
        var end_address = getCharFromNumber(values[0].length - 1 + add_col) + (values.length + row_offset - 1).toString();
        var rangeAddress = startCell + ":" + end_address;
        //console.log(rangeAddress)
        var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
        var range = worksheet.getRange(rangeAddress);

        var range_all = worksheet.getRange();
        var used_range = range_all.getUsedRange();

        used_range.clear();
        range.values = values;
        range.load('text');
        return ctx.sync().then(function() {
            // console.log(range.text);
            //window.location = "mac_start.html";
        });
    }).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });

}
