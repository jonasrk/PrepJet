(function () {
    'use strict';

    var worksheets = null;

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {
            app.initialize();

            $('#get-data-from-selection').click(getDataFromSelection);

            $('#step2').hide();
            $('#step3').hide();

            populateDropdowns();

            $('#bt_step2').click(step2ButtonClicked);
            $('#bt_step3').click(step3ButtonClicked);
            $('#bt_apply').click(applyButtonClicked);


        });
    };

    function step2ButtonClicked() {

        $('#step1').hide();
        $('#step2').show();
        $('#step3').hide();

        var selected_table2 = document.getElementById('table2_options').value; // TODO better reference by ID than name

        Excel.run(function (ctx) {
            var worksheet = ctx.workbook.worksheets.getItem(selected_table2);

            var rangeAddress = "A:Z"; // TODO Z is not the maximum
            var range_all = worksheet.getRange(rangeAddress);
            var range = range_all.getUsedRange();

            range.load('address');
            range.load('text');
            return ctx.sync().then(function() {
                for (var i = 0; i < range.text[0].length; i++) {

                    var el = document.createElement("div");
                    el.className = "ms-ChoiceField";
                    var el2 =  document.createElement("input");
                    el2.className = "ms-ChoiceField-input";
                    el2.id = "demo-checkbox-unselected";
                    el2.setAttribute("type", "checkbox");
                    var el3 = document.createElement("label");
                    el3.setAttribute("for", "checkbox");
                    el3.className = "ms-ChoiceField-field";
                    var el4 = document.createElement("span");
                    el4.className = "ms-Label";
                    el4.textContent = range.text[0][i];

                    el.appendChild(el2);
                    el.appendChild(el3);
                    el.appendChild(el4);

                    document.getElementById("checkboxes_variables").appendChild(el);

                }

                // $(".ms-ChoiceField").ChoiceField();
            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function step3ButtonClicked() {
        $('#step1').hide();
        $('#step2').hide();
        $('#step3').show();

        var selected_table1 = document.getElementById('table1_options').value; // TODO better reference by ID than name
        var selected_table2 = document.getElementById('table2_options').value; // TODO better reference by ID than name

        Excel.run(function (ctx) {
            var worksheet = ctx.workbook.worksheets.getItem(selected_table1);

            var rangeAddress = "A:Z"; // TODO Z is not the maximum
            var range_all = worksheet.getRange(rangeAddress);
            var range = range_all.getUsedRange();

            range.load('address');
            range.load('text');
            return ctx.sync().then(function() {
                for (var i = 0; i < range.text[0].length; i++) {

                    var el = document.createElement("option");
                    el.value = range.text[0][i];
                    el.textContent = range.text[0][i];
                    document.getElementById("reference_column_ckeckboxes_1").appendChild(el);

                }

                $(".reference_column_ckeckboxes_1").Dropdown();
            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

        Excel.run(function (ctx) {
            var worksheet = ctx.workbook.worksheets.getItem(selected_table2);

            var rangeAddress = "A:Z"; // TODO Z is not the maximum
            var range_all = worksheet.getRange(rangeAddress);
            var range = range_all.getUsedRange();

            range.load('address');
            range.load('text');
            return ctx.sync().then(function() {
                for (var i = 0; i < range.text[0].length; i++) {

                    var el = document.createElement("option");
                    el.value = range.text[0][i];
                    el.textContent = range.text[0][i];
                    document.getElementById("reference_column_ckeckboxes_2").appendChild(el);

                }

                $(".reference_column_ckeckboxes_2").Dropdown();
            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function applyButtonClicked() {
        $('#step1').show();
        $('#step2').hide();
        $('#step3').hide();

        var selected_table1 = document.getElementById('table1_options').value; // TODO better reference by ID than name
        var selected_table2 = document.getElementById('table2_options').value; // TODO better reference by ID than name

        console.log(document.getElementById('reference_column_ckeckboxes_1').value);
        console.log(document.getElementById('reference_column_ckeckboxes_2').value);

        Excel.run(function (ctx) {
            var worksheet = ctx.workbook.worksheets.getItem(selected_table2);

            var range_all = worksheet.getRange();
            var range = range_all.getUsedRange();

            range.load('address');
            range.load('text');

            var worksheet_adding_to = ctx.workbook.worksheets.getItem(selected_table1);

            var range_all_adding_to = worksheet_adding_to.getRange();
            var range_adding_to = range_all_adding_to.getUsedRange();

            range_adding_to.load('address');
            range_adding_to.load('text');


            return ctx.sync().then(function() {

                for (var i = 1; i < range.text.length; i++) {

                    console.log(range.text[i][0]); // TODO do not hardcode column

                    for (var i = 1; i < range_adding_to.text.length; i++) {

                        // TODO do not hardcode column

                        if (range_adding_to.text[i][0] == range.text[i][0]) {

                            console.log('found Match!');
                            console.log(range_adding_to.text[i][1]);
                            console.log(range.text[i][1]);
                            addContentToWorksheet(worksheet_adding_to, "J"+i , range.text[i][1])

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

    // Helper function to add and format content in the workbook
    function addContentToWorksheet(sheetObject, rangeAddress, displayText) {

        // Format differently by the type of content
        var range = sheetObject.getRange(rangeAddress);
        range.values = displayText;
        // range.format.font.name = "Corbel";
        // range.format.font.size = 30;
        // range.format.font.color = "white";
        range.merge();
    }

    function populateDropdowns() {

        var allworksheets = [];

        Excel.run(function (ctx) {
            worksheets = ctx.workbook.worksheets;
            worksheets.load('items');
            return ctx.sync().then(function () {
                for (var i = 0; i < worksheets.items.length; i++) {
                    worksheets.items[i].load('name');
                    worksheets.items[i].load('index');
                    ctx.sync().then(function (i) {

                        var this_i = i;

                        return function () {
                            allworksheets.push(worksheets.items[this_i].name);

                            if (this_i == worksheets.items.length - 1) {

                                for (var i = 0; i < allworksheets.length; i++) {
                                    var opt = allworksheets[i];
                                    var el = document.createElement("option");
                                    el.textContent = opt;
                                    el.value = opt;
                                    document.getElementById("table1_options").appendChild(el);
                                    var el = document.createElement("option");
                                    el.textContent = opt;
                                    el.value = opt;
                                    document.getElementById("table2_options").appendChild(el);
                                }

                                $(".dropdown_table").Dropdown();

                            }
                        }

                    }(i));
                }

            });

        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });


    }

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    app.showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    app.showNotification('Error:', result.error.message);
                }
            }
        );
    }

})();
