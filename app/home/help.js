(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            app.initialize();
            $('#addColumn').hide();
            $('#splitValues').hide();
            $('#extractValues').hide();
            $('#trimSpaces').hide();
            $('#detectDuplicates').hide();
            $('#detectOutlier').hide();
            $('#validationRule').hide();

            $('#goToHomeLabel').click(home);
            $('#linkAddColumn').click(addColumn);
            $('#linkSplitValues').click(splitValues);
            $('#linkExtractValues').click(extractValues);
            $('#linkTrimSpaces').click(trimSpaces);
            $('#linkDetectDuplicates').click(detectDuplicates);
            $('#linkDetectOutlier').click(detectOutlier);
            $('#linkValidationRule').click(validationRule);

        });
    };


    function home() {

        $('#firstpage').show();
        $('#addColumn').hide();
        $('#splitValues').hide();
        $('#extractValues').hide();
        $('#trimSpaces').hide();
        $('#detectDuplicates').hide();
        $('#detectOutlier').hide();
        $('#validationRule').hide();

        Excel.run(function (ctx) {

            return ctx.sync().then(function() {

            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }


    function addColumn() {

        $('#firstpage').hide();
        $('#addColumn').show();
        $('#splitValues').hide();
        $('#extractValues').hide();
        $('#trimSpaces').hide();
        $('#detectDuplicates').hide();
        $('#detectOutlier').hide();
        $('#validationRule').hide();

        Excel.run(function (ctx) {

            return ctx.sync().then(function() {

            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }


    function splitValues() {

        $('#firstpage').hide();
        $('#addColumn').hide();
        $('#splitValues').show();
        $('#extractValues').hide();
        $('#trimSpaces').hide();
        $('#detectDuplicates').hide();
        $('#detectOutlier').hide();
        $('#validationRule').hide();

        Excel.run(function (ctx) {

            return ctx.sync().then(function() {

            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }



    function extractValues() {

        $('#firstpage').hide();
        $('#addColumn').hide();
        $('#splitValues').hide();
        $('#extractValues').show();
        $('#trimSpaces').hide();
        $('#detectDuplicates').hide();
        $('#detectOutlier').hide();
        $('#validationRule').hide();

        Excel.run(function (ctx) {

            return ctx.sync().then(function() {

            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }




    function trimSpaces() {

        $('#firstpage').hide();
        $('#addColumn').hide();
        $('#splitValues').hide();
        $('#extractValues').hide();
        $('#trimSpaces').show();
        $('#detectDuplicates').hide();
        $('#detectOutlier').hide();
        $('#validationRule').hide();

        Excel.run(function (ctx) {

            return ctx.sync().then(function() {

            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }




    function detectDuplicates() {

        $('#firstpage').hide();
        $('#addColumn').hide();
        $('#splitValues').hide();
        $('#extractValues').hide();
        $('#trimSpaces').hide();
        $('#detectDuplicates').show();
        $('#detectOutlier').hide();
        $('#validationRule').hide();

        Excel.run(function (ctx) {

            return ctx.sync().then(function() {

            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }




    function detectOutlier() {

        $('#firstpage').hide();
        $('#addColumn').hide();
        $('#splitValues').hide();
        $('#extractValues').hide();
        $('#trimSpaces').hide();
        $('#detectDuplicates').hide();
        $('#detectOutlier').show();
        $('#validationRule').hide();

        Excel.run(function (ctx) {

            return ctx.sync().then(function() {

            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }




    function validationRule() {

        $('#firstpage').hide();
        $('#addColumn').hide();
        $('#splitValues').hide();
        $('#extractValues').hide();
        $('#trimSpaces').hide();
        $('#detectDuplicates').hide();
        $('#detectOutlier').hide();
        $('#validationRule').show();

        Excel.run(function (ctx) {

            return ctx.sync().then(function() {

            });

        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }


})();