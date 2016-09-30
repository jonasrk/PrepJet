function redirectHome() {
    window.location = "mac_start.html";
}

var activeSelection = 1;
function setFocus(activeID) {
    activeSelection = activeID;
}

(function () {
    // 'use strict';
    var fixCount = 1;
    var typeCount = 1;

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            Office.context.document.settings.set('same_header_trim', false);
            Office.context.document.settings.set('on_second_page', false);
            Office.context.document.settings.set('last_clicked_function', "temp_feature.html");
            if (Office.context.document.settings.get('prepjet_loaded_before') == null) {
                Office.context.document.settings.set('backup_sheet_count', 1);
                Office.context.document.settings.set('prepjet_loaded_before', true);
                Office.context.document.settings.saveAsync();
                window.location = "intro.html";
            }


            app.initialize();

            $('#step2').hide();
            $('#bt_remove').hide();
            $('#bt2_remove').hide();

            $('#helpCallout').hide();
            $('#check_template').click(compareTemplate);
            $('#homeButton').click(redirectHome);
            $('#continue1').click(showStep2);
            $('#bt_more').click(addContentTextField);
            $('#bt_remove').click(removeContentField);


            document.getElementById("refresh_icon").onclick = function () {
                window.location = "temp_feature.html";
            }


            //hide result message
            document.getElementById("resultClose").onclick = function () {
                document.getElementById('resultDialog').style.visibility = 'hidden';
                window.location = "temp_feature.html";
            }
            document.getElementById("resultOk").onclick = function () {
                document.getElementById('resultDialog').style.visibility = 'hidden';
                window.location = "temp_feature.html";
            }

            Office.context.document.addHandlerAsync("documentSelectionChanged", myHandler, function(result){}
            );
            // Event handler function.
            function myHandler(eventArgs){
                Excel.run(function (ctx) {
                    var selectedRange = ctx.workbook.getSelectedRange();
                    selectedRange.load('address');
                    return ctx.sync().then(function () {
                        if (Office.context.document.settings.get('on_second_page') == false) {
                            writeContent(selectedRange.address, activeSelection);
                        } else {
                            writeType(selectedRange.address, activeSelection);
                        }
                    });
                });
            }
            // Function that writes to a div with id='message' on the page.
            function writeContent(message, selection){
                document.getElementById('fixedContentInput' + selection).value = message;
            }
            function writeType(message, selection){
                document.getElementById('fixedTypeInput' + selection).value = message;
            }

        });
    };

    function addContentTextField() {

        fixCount += 1;

        var div = document.createElement("div");
        div.className = "ms-TextField ms-TextField--placeholder";
        div.id = "fixedContent" + fixCount;

        var label = document.createElement("label");
        label.innerHTML = "Select " + fixCount + ". Range:";

        var input = document.createElement("input");
        input.id = "fixedContentInput" + fixCount;
        input.className = "ms-TextField-field";
        input.addEventListener = ('onfocus', setFocus(fixCount));

        div.appendChild(label);
        div.appendChild(input);

        document.getElementById("contentDiv").appendChild(div);
        $('#bt_remove').show();

    }

    function removeContentField() {
        var parent = document.getElementById('contentDiv');
        var child = document.getElementById('fixedContent' + fixCount);
        parent.removeChild(child);
        fixCount -= 1;
        if (fixCount < 2) {
            $('#bt_remove').hide();
        }
    }



    function showStep2() {

        $('#step2').show();
        $('#step1').hide();

        Office.context.document.settings.set('on_second_page', true);

        function addField() {
            typeCount += 1;
            addTextField(typeCount);
            $('#bt2_remove').show();
        }

        function addTextField(id) {

            var div = document.createElement("div");
            div.className = "ms-TextField ms-TextField--placeholder";
            div.id = "fixedType" + id;

            var label = document.createElement("label");
            label.innerHTML = "Select " + typeCount + ". Range";

            var input = document.createElement("input");
            input.id = "fixedTypeInput" + id;
            input.className = "ms-TextField-field";
            input.addEventListener = ('onfocus', setFocus(typeCount));

            div.appendChild(label);
            div.appendChild(input);

            document.getElementById("typeDiv").appendChild(div);
        }

        function removeTypeField() {
            var parent = document.getElementById('typeDiv');
            var child = document.getElementById('fixedType' + typeCount)
            parent.removeChild(child);
            typeCount -= 1;
            if (typeCount <= 1) {
                $('#bt2_remove').hide();
            }
        }

        $("#bt2_more").unbind('click');
        $('#bt2_more').click(addField);
        $('#bt2_remove').click(removeTypeField);
    }



    function fillColumn(){

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange(true);

            range.load('text');
            worksheet.load('name');

            return ctx.sync().then(function() {

            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

    }


    function compareTemplate() {

        Excel.run(function (ctx) {

            var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange(true);
            var firstCell = range.getColumn(0);
            var firstCol = firstCell.getEntireColumn();
            var tmpRow = range.getRow(0);
            var firstRow = tmpRow.getEntireRow();

            //get used range in active Sheet
            range.load('text');
            worksheet.load('name');
            firstRow.load('address');
            firstCol.load('address');

            return ctx.sync().then(function() {

                var header = 0;
                var act_worksheet = ctx.workbook.worksheets.getActiveWorksheet();
                var tmp_offset = firstCol.address;
                var col_offset = tmp_offset.substring(tmp_offset.indexOf("!") + 1, tmp_offset.indexOf(":"));
                var tmp_row = firstRow.address;
                var row_offset = Number(tmp_row.substring(tmp_row.indexOf("!") + 1, tmp_row.indexOf(":")));
                var add_col = getNumberFromChar(col_offset);
                var startCell = col_offset + row_offset;

                backupForUndo(range, startCell, add_col, row_offset);

                var fixedAddresses = [];
                for (var i = 0; i < fixCount; i++) {
                    var tmpAddress = document.getElementById('fixedContentInput' + (i + 1)).value;
                    tmpAddress = tmpAddress.substring(tmpAddress.indexOf("!") + 1);
                    fixedAddresses.push(tmpAddress);
                }

                var typeAddresses = [];
                for (var i = 0; i < typeCount; i++) {
                    var tmpAddress = document.getElementById('fixedTypeInput' + (i + 1)).value;
                    tmpAddress = tmpAddress.substring(tmpAddress.indexOf("!") + 1);
                    typeAddresses.push(tmpAddress);
                }

                console.log(typeAddresses);

                if (document.getElementById('createBackup').checked == true) {
                    var sheet_count = Office.context.document.settings.get('backup_sheet_count') + 1;
                    Office.context.document.settings.set('backup_sheet_count', sheet_count);
                    Office.context.document.settings.saveAsync();
                    var newName = worksheet.name + "(" + sheet_count + ")";
                    addBackupSheet(newName, startCell, add_col, row_offset, function() {
                        var txt = document.createElement("p");
                        txt.className = "ms-font-xs ms-embedded-dialog__content__text";
                        //txt.innerHTML = "PrepJet successfully removed all leading and trailing spaces in the " + checked_checkboxes.length + endString;
                        document.getElementById('resultText').appendChild(txt);
                        document.getElementById('resultDialog').style.visibility = 'visible';
                    });

                } else {
                    var txt = document.createElement("p");
                    txt.className = "ms-font-xs ms-embedded-dialog__content__text";
                    //txt.innerHTML = "PrepJet successfully removed all leading and trailing spaces in the " + checked_checkboxes.length + endString;
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