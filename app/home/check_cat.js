function redirectHome() {
    window.location = "mac_start.html";
}


(function () {
    var count_wrong_cats = 0;
    // 'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            Office.context.document.settings.set('same_header_split', false);
            Office.context.document.settings.set('more_option', false);
            Office.context.document.settings.set('last_clicked_function', "split_values.html");
            if (Office.context.document.settings.get('prepjet_loaded_before') == null) {
                Office.context.document.settings.set('backup_sheet_count', 1);
                Office.context.document.settings.set('prepjet_loaded_before', true);
                Office.context.document.settings.saveAsync();
                window.location = "intro.html";
            }

            app.initialize();
            fillColumn();

            $zopim(function() {
                $zopim.livechat.window.hide();
            });

            $('#step2').hide();
            $('#changeDialog').hide();

            $('#check_categories').click(checkCategories);
            $('#buttonOk').click(highlightHeader);
            $('#homeButton').click(redirectHome);
            $('#change_categories').click(changeCategories);


            // Hides the dialog.
            document.getElementById("buttonClose").onclick = function () {
                document.getElementById('showEmbeddedDialog').style.visibility = 'hidden';
            }

            document.getElementById("refresh_icon").onclick = function () {
                window.location = "check_cat.html";
            }

            document.getElementById("resultClose").onclick = function () {
                document.getElementById('resultDialog').style.visibility = 'hidden';
                $('#step1').hide();
                $('#step2').show();
                //window.location = "check_cat.html";
            }
            document.getElementById("resultOk").onclick = function () {
                document.getElementById('resultDialog').style.visibility = 'hidden';
                $('#step1').hide();
                $('#step2').show();
                //window.location = "check_cat.html";
            }


            /*Excel.run(function (ctx) {

                //var myBindings = Office.context.document.bindings;
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
                                    window.location = "split_values.html"
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
                    window.location = "split_values.html";
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
                            Office.context.document.settings.set('same_header_split', true);
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
                    document.getElementById("column_options").appendChild(el);
                }
                $(".dropdown_table_col").Dropdown();
                $("span.ms-Dropdown-title:empty").text(range.text[0][0]);
            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

    }


    function fillCategories(cat_object){

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('text');

            return ctx.sync().then(function() {

                for (var key in cat_object) {
                    var key_name = key + " (" + cat_object[key] + ")";
                    createTableRow(key_name);
                }

                function createTableRow(keyname) {

                    var trow = document.createElement("tr");
                    var tcol1 = document.createElement("td");
                    var tcol2 = document.createElement("td");
                    trow.appendChild(tcol1);
                    trow.appendChild(tcol2);
                    document.getElementById('checkboxes_categories').appendChild(trow);

                    var label = document.createElement("label");
                    label.id = "newLabel" + count_wrong_cats;
                    label.innerHTML = keyname;

                    tcol1.appendChild(label);

                    var textfield = document.createElement("div");
                    textfield.className = "ms-TextField";

                    var input = document.createElement("input");
                    input.id = "newCat" + count_wrong_cats;
                    input.className = "ms-TextField-field";
                    input.type = "text";
                    textfield.appendChild(input);

                    tcol2.appendChild(textfield);
                    count_wrong_cats += 1;
                }

            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

    }


    function changeCategories(){

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();
            var selected_identifier = document.getElementById('column_options').value;

            range.load('text');

            return ctx.sync().then(function() {

                var header = 0;
                for (var k = 0; k < range.text[0].length; k++){
                    if (selected_identifier == range.text[0][k] || selected_identifier == "Column " + getCharFromNumber(k)){
                        header = k;
                    }
                }

                for (var i = 0; i < count_wrong_cats; i++) {
                    var tmp_name = "newCat" + i;
                    var tmp_label = "newLabel" + i;
                    var newCatName = document.getElementById(tmp_name).value;
                    var oldCatName = document.getElementById(tmp_label).innerHTML;
                    oldCatName = oldCatName.substring(0, oldCatName.indexOf("(") - 1);
                    console.log(oldCatName)
                    if (newCatName != "") {
                        for (var k = 0; k < range.text.length; k++) {
                            if (range.text[k][header] == oldCatName) {
                                var insertAddress = getCharFromNumber(header) + (k + 1);
                                addContentToWorksheet(worksheet, insertAddress, newCatName)
                            }
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


    //function to split values in a column by a specified delimiter into different columns
    function checkCategories() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();
            var selected_identifier = document.getElementById('column_options').value;

            range.load('text');
            worksheet.load('name');


            return ctx.sync().then(function() {

                backupForUndo(range);

                //get column number which to split
                var header = 0;
                for (var k = 0; k < range.text[0].length; k++){
                    if (selected_identifier == range.text[0][k] || selected_identifier == "Column " + getCharFromNumber(k)){
                        header = k;
                    }
                }

                var categories = {};
                for(var i = 1; i < range.text.length; ++i) {
                    if(!categories[range.text[i][header]])
                        categories[range.text[i][header]] = 0;
                    ++categories[range.text[i][header]];
                }

                var count_categories = Object.keys(categories).length;
                var count_data_records = range.text.length;

                var value_count = [];
                for (var key in categories) {
                    value_count.push(categories[key]);
                }

                var keysSorted = sortobj(categories);
                fillCategories(keysSorted);

                function sortobj(obj) {
                    var keys=Object.keys(obj);
                    var kva= keys.map(function(k,i)
                    {
                        return [k,obj[k]];
                    });
                    kva.sort(function(a,b){
                        if(b[1]>a[1]) return -1;if(b[1]<a[1]) return 1;
                        return 0
                    });
                    var o={}
                    kva.forEach(function(a){ o[a[0]]=a[1]})
                    return o;
                }

                if (document.getElementById('createBackup').checked == true) {
                    var sheet_count = Office.context.document.settings.get('backup_sheet_count') + 1;
                    Office.context.document.settings.set('backup_sheet_count', sheet_count);
                    Office.context.document.settings.saveAsync();
                    var newName = worksheet.name + "(" + sheet_count + ")";
                    addBackupSheet(newName, function() {
                        var txt = document.createElement("p");
                        txt.className = "ms-font-xs ms-embedded-dialog__content__text";
                        txt.innerHTML = "PrepJet found " + count_categories + " categories."
                        document.getElementById('resultText').appendChild(txt);

                        document.getElementById('resultDialog').style.visibility = 'visible';
                    });
                }
                else {
                    var txt = document.createElement("p");
                    txt.className = "ms-font-xs ms-embedded-dialog__content__text";
                    txt.innerHTML = "PrepJet found " + count_categories + " categories."
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
