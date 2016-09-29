function redirectTrim() {
    window.location = "trim_spaces.html";
}

function redirectHarm() {
    window.location = "harmonize.html";
}

function redirectExtract() {
    window.location = "extract_values.html";
}

function redirectSplit() {
    window.location = "split_values.html";
}

function redirectIncon() {
    window.location = "inconsistency.html";
}

function redirectCustomIncon() {
    window.location = "custom_incon.html";
}

function redirectLookup() {
    window.location = "merge_columns.html";
}

function redirectDuplicates() {
    window.location = "duplicates.html";
}

function redirectValidation() {
    window.location = "validation.html";
}

function redirectHelp() {
    window.location = "help.html";
}

(function () {
    // 'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            app.initialize();

            $('#trimButton').click(redirectTrim);
            $('#harmonizeButton').click(redirectHarm);
            $('#extractButton').click(redirectExtract);
            $('#splitButton').click(redirectSplit);
            $('#lookupButton').click(redirectLookup);
            $('#inconButton').click(redirectIncon);
            $('#custominconButton').click(redirectCustomIncon);
            $('#duplicatesButton').click(redirectDuplicates);
            $('#validationButton').click(redirectValidation);
            $('#helpButton').click(redirectHelp);

        });
    };

})();