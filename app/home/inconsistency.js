(function () {
    // 'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            Office.context.document.settings.set('same_header_inconsistency', false);
            Office.context.document.settings.set('last_clicked_function', "inconsistency.html");
            if (Office.context.document.settings.get('prepjet_loaded_before') == null) {
                Office.context.document.settings.set('prepjet_loaded_before', true);
                Office.context.document.settings.saveAsync();
                window.location = "intro.html";
            }

            app.initialize();
            fillColumn();


            $('#inconsistency').click(inconsistencies);
            $('#checkbox_all').click(checkCheckbox);
            $('#buttonOk').click(highlightHeader);

            // Show and hide error message if columns have same header name
            document.getElementById("buttonClose").onclick = function () {
                document.getElementById('showEmbeddedDialog').style.visibility = 'hidden';
            }


            //show and hide help callout
            document.getElementById("help_icon").onclick = function () {
                document.getElementById('helpCallout').style.visibility = 'visible';
            }
            document.getElementById("closeCallout").onclick = function () {
                document.getElementById('helpCallout').style.visibility = 'hidden';
            }

            document.getElementById("refresh_icon").onclick = function () {
                window.location = "inconsistencies.html";
            }

            //hide result message
            document.getElementById("resultClose").onclick = function () {
                document.getElementById('resultDialog').style.visibility = 'hidden';
                window.location = "inconsistency.html";
            }
            document.getElementById("resultOk").onclick = function () {
                document.getElementById('resultDialog').style.visibility = 'hidden';
                window.location = "inconsistency.html";
            }


            /*Excel.run(function (ctx) {

                var myBindings = Office.context.document.bindings;
                var worksheetname = ctx.workbook.worksheets.getActiveWorksheet();

                worksheetname.load('name')

                return ctx.sync().then(function() {

                    Office.context.document.addHandlerAsync("documentSelectionChanged", myViewHandler, function(result){}
                    );

                    // Event handler function for changing the worksheet.
                    function myViewHandler(eventArgs){
                        Excel.run(function (ctx) {
                            var selectedSheet = ctx.workbook.worksheets.getActiveWorksheet();
                            selectedSheet.load('name');
                            return ctx.sync().then(function () {
                                if (selectedSheet.name != worksheetname.name) {
                                    window.location = "inconsistency.html"
                                }
                            });
                        });
                    }

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
                    window.location = "inconsistency.html";
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
            });*/

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


    function checkCheckbox() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('address');
            range.load('text');

            return ctx.sync().then(function() {
                if (document.getElementById('checkbox_all').checked == true) {
                    for (var i = 0; i < range.text[0].length; i++) {
                        if (range.text[0][i] != "") {
                            document.getElementById(range.text[0][i]).checked = true;
                        }
                        else {
                            document.getElementById("Column " + getCharFromNumber(i)).checked = true;
                        }
                    }
                }
                else {
                    for (var i = 0; i < range.text[0].length; i++) {
                        if (range.text[0][i] != "") {
                            document.getElementById(range.text[0][i]).checked = false;
                        }
                        else {
                            document.getElementById("Column " + getCharFromNumber(i)).checked = false;
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


    function fillColumn(){

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('text');

            return ctx.sync().then(function() {

                for (var run = 0; run < range.text[0].length - 1; run++) {
                    for (var run2 = run + 1; run2 < range.text[0].length; run2++) {
                        if (range.text[0][run] == range.text[0][run2] && range.text[0][run] != "") {
                            document.getElementById('showEmbeddedDialog').style.visibility = 'visible';
                            Office.context.document.settings.set('same_header_inconsistency', true);
                        }
                    }
                }

                for (var i = 0; i < range.text[0].length; i++) {
                    if (range.text[0][i] != ""){
                        addNewCheckboxToContainer (range.text[0][i], "column_checkbox" ,"checkboxes_columns");
                    }
                    else {
                        var colchar = getCharFromNumber(i);
                        addNewCheckboxToContainer ("Column " + colchar, "column_checkbox" ,"checkboxes_columns");
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


    function inconsistencies() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('text');
            range.load('valueTypes'); //does not know date type
            range.load('values');
            range.load('numberFormat');

            return ctx.sync().then(function() {

                var header = 0;
                var checked_checkboxes = getCheckedBoxes("column_checkbox");
                var check = [];
                var incon_counter = 0;

                for (var run = 0;run < checked_checkboxes.length; run++) {
                    check[run] = 0;
                    var type_maximum = 0;
                    for (var k = 0; k < range.text[0].length; k++) {
                        if (checked_checkboxes[run].id == range.text[0][k] || checked_checkboxes[run].id == "Column " + getCharFromNumber(k)){
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
                        tmp_address = getCharFromNumber(header) + (i + 1);
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
                                        incon_counter += 1;
                                    }
                                }
                            }
                            if (tmp2[i][1] == type_maximum && equal_type_check > 1) {
                                color = getRandomColor();
                                for (var k = 0; k < tmp_type.length; k++) {
                                    if (tmp2[i][0] == tmp_type[k][0]) {
                                        highlightCellInWorksheet(worksheet, tmp_type[k][1], color);
                                        incon_counter += 1;
                                    }
                                }
                            }
                        }
                    }

                }

                var txt = document.createElement("p");
                txt.className = "ms-font-xs ms-embedded-dialog__content__text";
                txt.innerHTML = "PrepJet found " + incon_counter + " inconsistent data entries in your worksheet."
                document.getElementById('resultText').appendChild(txt);

                document.getElementById('resultDialog').style.visibility = 'visible';

                //window.location = "inconsistency.html";
            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }


})();