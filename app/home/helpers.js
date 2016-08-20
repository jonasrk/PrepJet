/**
 * Created by jonas on 6/23/16.
 */

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

// Pass the checkbox name to the function
function getAllCheckBoxes(chkboxName) {
    var checkboxes = document.getElementsByName(chkboxName);
    var checkboxesChecked = [];
    // loop over them all
    for (var i=0; i<checkboxes.length; i++) {
        checkboxesChecked.push(checkboxes[i]);
        }
    // Return the array if it is non-empty, or null
    return checkboxesChecked.length > 0 ? checkboxesChecked : null;
}

// create a ms-ChoiceField html element (input + div + label + span) for every column in the selected table
function addNewCheckboxToContainer (id, name, container) {

    var el =  document.createElement("input");
    el.id = id;
    el.name = name;
    el.className = "ms-ChoiceField-input";
    el.setAttribute("type", "checkbox");

    var div = document.createElement("div");

    var label =  document.createElement("label");
    label.className = "ms-ChoiceField-field";
    label.setAttribute("for", id);

    var span =  document.createElement("span");
    span.className = "ms-Label";
    span.textContent = id;

    label.appendChild(span);
    div.appendChild(el);
    div.appendChild(label);

    document.getElementById(container).appendChild(div);

}


function getCharFromNumber (number) {

    if (number == 0) {
        return 'A';
    } else if (number == 1) {
        return 'B';
    } else if (number == 2) {
        return 'C';
    } else if (number == 3) {
        return 'D';
    } else if (number == 4) {
        return 'E';
    } else if (number == 5) {
        return 'F';
    } else if (number == 6) {
        return 'G';
    } else if (number == 7) {
        return 'H';
    } else if (number == 8) {
        return 'I';
    } else if (number == 9) {
        return 'J';
    } else if (number == 10) {
        return 'K';
    } else if (number == 11) {
        return 'L';
    } else if (number == 12) {
        return 'M';
    } else if (number == 13) {
        return 'N';
    } else if (number == 14) {
        return 'O';
    } else if (number == 15) {
        return 'P';
    } else if (number == 16) {
        return 'Q';
    } else if (number == 17) {
        return 'R';
    } else if (number == 18) {
        return 'S';
    } else if (number == 19) {
        return 'T';
    } else if (number == 20) {
        return 'U';
    } else if (number == 21) {
        return 'V';
    } else if (number == 22) {
        return 'W';
    } else if (number == 23) {
        return 'X';
    } else if (number == 24) {
        return 'Y';
    } else if (number == 25) {
        return 'Z';
    }

    if (number > 25) {
        return getCharFromNumber(Math.floor(number / 26) - 1) + getCharFromNumber(number % 26);
    }

}


function addBackupSheet(sheetName, callback) {
    Excel.run(function (ctx) {
        var wSheetName = sheetName;
        var worksheet = ctx.workbook.worksheets.add(wSheetName);
        worksheet.load('name');
        return ctx.sync().then(function() {
            addBackupContent(worksheet.name, callback);
        });
    }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
    });
}

function addBackupContent(sheetName, callback) {
    Excel.run(function (ctx) {
        var values = Office.context.document.settings.get('sheet_backup');
        var end_address = getCharFromNumber(values[0].length - 1) + (values.length).toString();
        var rangeAddress = "A1:" + end_address;
        var worksheet = ctx.workbook.worksheets.getItem(sheetName);
        var range = worksheet.getRange(rangeAddress);

        range.values = values;
        range.load('text');

        return ctx.sync().then(function() {
            callback();
        });
    }).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}

//inserts whole column
function addContentNew(sheetObject, rangeAddress, displayText, callback) {
    Excel.run(function (ctx) {
        var range = ctx.workbook.worksheets.getItem(sheetObject).getRange(rangeAddress);
        range.values = displayText;
        range.load('text');
        return ctx.sync().then(function() {
            callback();
        });
    }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
    });
}


// Helper function to add and format content (in single cell) in the workbook
function addContentToWorksheet(sheetObject, rangeAddress, displayText) {
    var range = sheetObject.getRange(rangeAddress);
    range.values = displayText;
    range.merge();
}

//color font in workbook
function highlightContentInWorksheet(sheetObject, rangeAddress, color) {
    var range = sheetObject.getRange(rangeAddress);
    range.format.font.color = color;
    range.merge();
}

function highlightContentNew(sheetObject, rangeAddress, color, callback) {
    Excel.run(function (ctx) {
        var range = ctx.workbook.worksheets.getItem(sheetObject).getRange(rangeAddress);
        range.format.font.color = color;
        return ctx.sync();
            callback();
    }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
    });
}

//color cell background
function highlightCellNew(sheetObject, rangeAddress, color, callback) {
    Excel.run(function (ctx) {
        var range = ctx.workbook.worksheets.getItem(sheetObject).getRange(rangeAddress);
        range.format.fill.color = color;
        return ctx.sync();
            callback();
    }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
    });
}

//create a random color to highlight font
function getRandomColor() {
    var letters = '0123456789ABCDEF'.split('');
    var color = '#';
    for (var i = 0; i < 6; i++ ) {
        color += letters[Math.floor(Math.random() * 16)];
    }
    return color;
}


function backupForUndo(this_range){

    Office.context.document.settings.set('sheet_backup', this_range.text);
    Office.context.document.settings.saveAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            console.log('Settings save failed. Error: ' + asyncResult.error.message);
        } else {
            console.log('Settings saved.');
            console.log(Office.context.document.settings.get('sheet_backup'));
        }
    });

}


function getColumn(sheet, colNumber, callback){

    Excel.run(function (ctx) {

        var worksheet = ctx.workbook.worksheets.getItem(sheet);
        var range_all = worksheet.getRange();
        var range = range_all.getUsedRange(true).getColumn(colNumber);

        range.load('text');
        range.load('address');

        return ctx.sync().then(function() {
            var tmp_columns = range.address;
            var column_address = tmp_columns.substring(tmp_columns.indexOf("!") + 1);
            callback(column_address);
        });

    }).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });

}


function detectIE() {
    var ua = window.navigator.userAgent;

    var msie = ua.indexOf('MSIE ');
    if (msie > 0) {
        // IE 10 or older => return version number
        window.location = 'error.html';
    }
}


//run on import:

detectIE();