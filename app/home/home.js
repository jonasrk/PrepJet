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
    }

    function populateDropdowns() {

        var allworksheets = [];

        Excel.run(function (ctx) {
            worksheets = ctx.workbook.worksheets;
            worksheets.load('items');
            return ctx.sync().then(function () {
                // console.log("### worksheets.items.length: " + worksheets.items.length);
                for (var i = 0; i < worksheets.items.length; i++) {
                    // console.log("### Loop iteration: " + i);
                    // console.log(worksheets.items[i]);
                    worksheets.items[i].load('name');
                    worksheets.items[i].load('index');
                    ctx.sync().then(function (i) {

                        var this_i = i;
                        // console.log("### this_i: " + this_i);

                        return function () {
                            // console.log(worksheets);
                            // console.log(worksheets.items);
                            // console.log(this_i);
                            // console.log(worksheets.items[this_i]);
                            // console.log(worksheets.items[this_i].name);
                            allworksheets.push(worksheets.items[this_i].name);
                            // console.log(worksheets.items[this_i].index);
                            // console.log(allworksheets);

                            if (this_i == worksheets.items.length - 1) {

                                // console.log(allworksheets);

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
