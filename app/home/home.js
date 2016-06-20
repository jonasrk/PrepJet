(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {
            app.initialize();

            $('#get-data-from-selection').click(getDataFromSelection);

            $('#step2').hide();
            $('#step3').hide();

            $('#bt_step2').click(step2ButtonClicked);
            $('#bt_step3').click(step3ButtonClicked);
            $('#bt_apply').click(applyButtonClicked);

            populateDropdowns();
        });
    };

    function step2ButtonClicked() {
        $('#step1').hide();
        $('#step2').show();
        $('#step3').hide();
    }

    function step3ButtonClicked() {
        $('#step1').hide();
        $('#step2').hide();
        $('#step3').show();
    }

    function applyButtonClicked() {
        $('#step1').show();
        $('#step2').hide();
        $('#step3').hide();
    }

    function populateDropdowns() {
        var all_worksheets = [];

        Excel.run(function (ctx) {
            var worksheets = ctx.workbook.worksheets;
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
                            all_worksheets.push(worksheets.items[this_i].name);
                            // console.log(worksheets.items[this_i].index);
                            // console.log(all_worksheets);

                            if (this_i == worksheets.items.length - 1) {

                                // console.log(all_worksheets);

                                for (var i = 0; i < all_worksheets.length; i++) {
                                    var opt = all_worksheets[i];
                                    var el = document.createElement("option");
                                    el.textContent = opt;
                                    el.value = opt;
                                    document.getElementById("table1_options").appendChild(el);
                                    var el = document.createElement("option");
                                    el.textContent = opt;
                                    el.value = opt;
                                    document.getElementById("table2_options").appendChild(el);
                                }

                                $(".ms-Dropdown").Dropdown();

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
