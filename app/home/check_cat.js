function redirectHome() {
    window.location = "mac_start.html";
}


(function () {
    var count_wrong_cats = 0;
    var count_corr_cats = 0;
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
            $('#showAll').click(showAllCats);


            //Hide and show help dialog
            document.getElementById("help_icon").onclick = function () {
                document.getElementById('helpCallout').style.visibility = 'visible';
            }
            document.getElementById("closeCallout").onclick = function () {
                document.getElementById('helpCallout').style.visibility = 'hidden';
            }

            // Hides the dialog for double column names.
            document.getElementById("buttonClose").onclick = function () {
                document.getElementById('showEmbeddedDialog').style.visibility = 'hidden';
            }

            //refresh side pane window
            document.getElementById("refresh_icon").onclick = function () {
                window.location = "check_cat.html";
            }

            //Close result window and load page 2
            document.getElementById("resultClose").onclick = function () {
                document.getElementById('resultDialog').style.visibility = 'hidden';
                $('#step1').hide();
                document.getElementById('step2').style.visibility = 'visible';
                document.getElementById('step2').style.display = 'block';
            }
            document.getElementById("resultOk").onclick = function () {
                document.getElementById('resultDialog').style.visibility = 'hidden';
                $('#step1').hide();
                document.getElementById('step2').style.visibility = 'visible';
                document.getElementById('step2').style.display = 'block';
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


    function fillCategories(cat_object, trstyle, counter){

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('text');

            return ctx.sync().then(function() {

                for (var key in cat_object) {
                    var key_name = key + " (" + cat_object[key] + ")";
                    createTableRow(key_name, counter);
                    counter -= 1;
                }

                function createTableRow(keyname) {

                    var trow = document.createElement("tr");
                    trow.id = "newRow" + counter;
                    var tcol1 = document.createElement("td");
                    var tcol2 = document.createElement("td");
                    trow.appendChild(tcol1);
                    trow.appendChild(tcol2);
                    document.getElementById('checkboxes_categories').appendChild(trow);

                    var label = document.createElement("label");
                    label.id = "newLabel" + counter;
                    label.innerHTML = keyname;

                    tcol1.appendChild(label);

                    var textfield = document.createElement("div");
                    textfield.className = "ms-TextField";

                    var input = document.createElement("input");
                    input.id = "newCat" + counter;
                    input.className = "ms-TextField-field";
                    input.type = "text";
                    textfield.appendChild(input);

                    tcol2.appendChild(textfield);

                    trow.style.visibility = trstyle;
                    if (trstyle == "hidden") {
                        trow.style.display = 'none';
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

                for (var i = count_wrong_cats; i > 0; i--) {
                    var tmp_name = "newCat" + i;
                    var tmp_label = "newLabel" + i;
                    var newCatName = document.getElementById(tmp_name).value;
                    var oldCatName = document.getElementById(tmp_label).innerHTML;
                    oldCatName = oldCatName.substring(0, oldCatName.indexOf("(") - 1);

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


    function showAllCats() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('text');

            return ctx.sync().then(function() {

                if (document.getElementById('showAll').checked == true) {
                    for (var i = (count_corr_cats + count_wrong_cats); i > count_wrong_cats ; i--) {
                        var tmp_name = "newRow" + i;
                        document.getElementById(tmp_name).style.visibility = "visible";
                        document.getElementById(tmp_name).style.display = "table-row";
                    }
                }
                else {
                    for (var i = (count_corr_cats + count_wrong_cats); i > count_wrong_cats ; i--) {
                        var tmp_name = "newRow" + i;
                        document.getElementById(tmp_name).style.visibility = "hidden";
                        document.getElementById(tmp_name).style.display = "none";
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
                //var count_suspicious = 0;

                var keysSorted = sortobj(categories);
                fillCategories(keysSorted.suspCat, "visible", count_wrong_cats);
                fillCategories(keysSorted.correctCat, "hidden", (count_wrong_cats + count_corr_cats));

                function sortobj(obj) {
                    var keys=Object.keys(obj);
                    var kva= keys.map(function(k,i) {
                        return [k,obj[k]];
                    });
                    kva.sort(function(a,b){
                        if(b[1]>a[1]) return -1;if(b[1]<a[1]) return 1;
                        return 0
                    });
                    var suspCat = {}
                    var correctCat = {}
                    kva.forEach(function(a) {
                        if (a[1] < 0.1 * count_data_records) {
                            suspCat[a[0]] = a[1]
                            //count_suspicious += 1;
                            count_wrong_cats += 1;
                        }
                        else {
                            correctCat[a[0]] = a[1];
                            count_corr_cats += 1;

                        }
                    })
                    return {suspCat: suspCat, correctCat: correctCat};
                }

                if (document.getElementById('createBackup').checked == true) {
                    var sheet_count = Office.context.document.settings.get('backup_sheet_count') + 1;
                    Office.context.document.settings.set('backup_sheet_count', sheet_count);
                    Office.context.document.settings.saveAsync();
                    var newName = worksheet.name + "(" + sheet_count + ")";
                    addBackupSheet(newName, function() {
                        var txt = document.createElement("p");
                        txt.className = "ms-font-xs ms-embedded-dialog__content__text";
                        txt.innerHTML = "PrepJet found " + count_categories + " categories of which " + count_wrong_cats + " are suspicious."
                        document.getElementById('resultText').appendChild(txt);

                        document.getElementById('resultDialog').style.visibility = 'visible';
                    });
                }
                else {
                    var txt = document.createElement("p");
                    txt.className = "ms-font-xs ms-embedded-dialog__content__text";
                    txt.innerHTML = "PrepJet found " + count_categories + " categories of which " + count_wrong_cats + " are suspicious."
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
