function redirectHome() {
    window.location = "mac_start.html";
}

function resultClose() {
    document.getElementById('resultDialog').style.visibility = 'hidden';
    window.location = "mac_start.html";
}

function resultOK() {
    document.getElementById('resultDialog').style.visibility = 'hidden';
    window.location = "mac_start.html";
}

(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            app.initialize();

            jQuery('#resultOk').click(resultOK);
            jQuery('#resultClose').click(resultClose);
            jQuery('#homeButton').click(redirectHome);
            jQuery('#inconsistency').click(checkIncon);

        });
    };


    function getSelectedData(callback) {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Matrix, { valueFormat: Office.ValueFormat.Formatted },
        function (result) {
            if (result.status == "succeeded") {
                callback(result.value);
            }
            else {
                console.log("error");
            }
        });
    }

    function getDataType(item) {
        var datatype = "string"//typeof item;
        return datatype;
    }


    function checkIncon() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function(result){
                getSelectedData(function(result){

                    if (result != null) {
                        var countIncon = 0;
                        var type_array = result.map(function (item) {
                            return item.map(function (item) {
                                if (item) {
                                    var itemType = getDataType(item);
                                    /*if (item != newitem) {
                                        countIncon++;
                                    }*/
                                    return itemType;
                                }
                            });
                        });
                    }

                    Office.context.document.setSelectedDataAsync(type_array, { valueFormat: Office.ValueFormat.Formatted }, function(result){
                        if (result.status == "succeeded") {
                            var txt = document.createElement("p");
                            txt.className = "ms-font-xs ms-embedded-dialog__content__text";
                            txt.innerHTML = "PrepJet found " + countIncon + " entries with inconsistent data type."
                            document.getElementById('resultText').appendChild(txt);
                            document.getElementById('resultDialog').style.visibility = 'visible';
                        } else {
                            console.log("An error occured. Please select a range and try again.");
                        }
                    });

                });
        });
    }



    /*function inconsistencies() {

        Excel.run(function (ctx) {


            return ctx.sync().then(function() {

                var header = 0;
                var check = [];
                var incon_counter = 0;

                for (var run = 0;run < checked_checkboxes.length; run++) {
                    check[run] = 0;
                    var type_maximum = 0;
                    for (var k = 0; k < range.text[0].length; k++) {
                        if (checked_checkboxes[run].id == range.text[0][k] || checked_checkboxes[run].id == "Column " + getCharFromNumber(k + add_col)){
                            header = k;
                            break;
                        }
                    }

                    var tmp_type = [];
                    var tmp_address = "";
                    var rangeType = [];
                    var tmpUniqueTypes = [];
                    for (var i = 1; i < range.text.length; i++) {
                        var rangeType = [];
                        tmp_address = getCharFromNumber(header + add_col) + (i + row_offset);
                        if (range.valueTypes[i][header] == "Double") {
                            if (range.numberFormat[i][header] != "General") {
                                rangeType.push("Date");
                            }
                            else {
                                rangeType.push(range.valueTypes[i][header]);
                            }
                        }
                        else {
                            rangeType.push(range.valueTypes[i][header]);
                        }

                        if (i == 1) {
                            tmpUniqueTypes.push(rangeType[0]);
                        }
                        rangeType.push(tmp_address);
                        tmp_type.push(rangeType);

                        if (i > 1 && (tmp_type[i - 1][0] != tmp_type[i - 2][0])) {
                            check[run] = 1;
                            var test_unique = 0
                            for (var k = 0; k < tmpUniqueTypes.length; k++) {
                                if (tmpUniqueTypes[k] != tmp_type[i - 1][0]) {
                                    test_unique = test_unique + 1;
                                }
                            }
                            if (test_unique >= tmpUniqueTypes.length) {
                                tmpUniqueTypes.push(tmp_type[i - 1][0]);
                            }
                        }
                    }

                    if (check[run] == 1) {
                        var tmp2 = [];
                        for (var j = 0; j < tmpUniqueTypes.length; j++) {
                            var type_counter = 0;
                            var tmp1 = [];
                            for (var i = 0; i < tmp_type.length; i++) {
                                if (tmp_type[i][0] == tmpUniqueTypes[j]) {
                                    type_counter = type_counter + 1;
                                }
                            }

                            if (type_maximum < type_counter) {
                                type_maximum = type_counter;
                            }
                            tmp1.push(tmpUniqueTypes[j]);
                            tmp1.push(type_counter);
                            tmp2.push(tmp1);
                        }

                        var equal_type_check = 0;
                        var empty_check = 0;
                        for (var i = 0; i < tmp2.length; i++) {
                            if (tmp2[i][1] == type_maximum) {
                                if (tmp2[i][0] == "Empty") {
                                    var t1 = tmp2.slice(0, i);
                                    var t2 = tmp2.slice(i + 1);
                                    tmp2 = t1.concat(t2);
                                    empty_check = 1;
                                }
                                equal_type_check += 1;
                            }
                        }

                        if (empty_check == 1) {
                            type_maximum = 0;
                            for (var i = 0; i < tmp2.length; i++) {
                                if (type_maximum < tmp2[i][1]) {
                                    type_maximum = tmp2[i][1];
                                }
                            } //todo again check whether new most frequent type occurs twice
                        }

                        equal_type_check = 0;
                        for (var i = 0; i < tmp2.length; i++) {
                            if (tmp2[i][1] == type_maximum) {
                                equal_type_check += 1;
                            }
                        }

                        var color = "#EA7F04";
                        for (var i = 0; i < tmp2.length; i++) {
                            if (tmp2[i][1] < type_maximum) {
                                for (var k = 0; k < tmp_type.length; k++) {
                                    if (tmp2[i][0] == tmp_type[k][0]) {
                                        highlightCellInWorksheet(worksheet, tmp_type[k][1], color);
                                        //highlightCellNew(worksheet.name, tmp_type[k][1], color, function () {});
                                        incon_counter += 1;
                                    }
                                }
                            }
                            if (tmp2[i][1] == type_maximum && equal_type_check > 1) {
                                color = getRandomColor();
                                for (var k = 0; k < tmp_type.length; k++) {
                                    if (tmp2[i][0] == tmp_type[k][0]) {
                                        highlightCellInWorksheet(worksheet, tmp_type[k][1], color);
                                        //highlightCellNew(worksheet.name, tmp_type[k][1], color, function () {});
                                        incon_counter += 1;
                                    }
                                }
                            }
                        }
                    }

                }


                if (document.getElementById('createBackup').checked == true) {
                    var sheet_count = Office.context.document.settings.get('backup_sheet_count') + 1;
                    Office.context.document.settings.set('backup_sheet_count', sheet_count);
                    Office.context.document.settings.saveAsync();
                    var newName = worksheet.name + "(" + sheet_count + ")";
                    var backup_promise = new Promise(
                        function(resolve, reject) {
                                resolve(addBackupSheet(newName, startCell, add_col, row_offset));
                        }
                    );

                    backup_promise.then(
                        function() {
                            var txt = document.createElement("p");
                            txt.className = "ms-font-xs ms-embedded-dialog__content__text";
                            txt.innerHTML = "PrepJet found " + incon_counter + " inconsistent data entries in your worksheet."
                            document.getElementById('resultText').appendChild(txt);

                            document.getElementById('resultDialog').style.visibility = 'visible';
                        })
                    .catch(
                        function(reason) {
                            console.log('Handle rejected promise ('+reason+') here.');
                        });
                }
                else {
                    var txt = document.createElement("p");
                    txt.className = "ms-font-xs ms-embedded-dialog__content__text";
                    txt.innerHTML = "PrepJet found " + incon_counter + " inconsistent data entries in your worksheet."
                    document.getElementById('resultText').appendChild(txt);

                    document.getElementById('resultDialog').style.visibility = 'visible';
                }


                //window.location = "inconsistency.html";
            });

        });
    }*/


})();