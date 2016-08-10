(function () {
    // 'use strict';

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
            $('#harmonize').hide();
            $('#undoHelp').hide();
            $('#tableHeader').hide();

            $('#splitValHome').click(home);
            $('#addColHome').click(home);
            $('#extractValHome').click(home);
            $('#trimSpacesHome').click(home);
            $('#detDupHome').click(home);
            $('#detOutHome').click(home);
            $('#valRuleHome').click(home);
            $('#harmonizeHome').click(home);
            $('#undoHelpHome').click(home);
            $('#tableHeaderHome').click(home);

            $('#linkAddColumn').click(addColumn);
            $('#linkSplitValues').click(splitValues);
            $('#linkExtractValues').click(extractValues);
            $('#linkTrimSpaces').click(trimSpaces);
            $('#linkDetectDuplicates').click(detectDuplicates);
            $('#linkDetectOutlier').click(detectOutlier);
            $('#linkValidationRule').click(validationRule);
            $('#linkharmonize').click(harmonizeColumn);
            $('#linkUndoHelp').click(undoHelp);
            $('#linkTableHeaderHelp').click(tableHeaderHelp);

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
        $('#harmonize').hide();
        $('#undoHelp').hide();
        $('#tableHeaderHelp').hide();

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
        $('#harmonize').hide();
        $('#undoHelp').hide();
        $('#tableHeaderHelp').hide();

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
        $('#harmonize').hide();
        $('#undoHelp').hide();
        $('#tableHeaderHelp').hide();

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
        $('#harmonize').hide();
        $('#undoHelp').hide();
        $('#tableHeaderHelp').hide();

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
        $('#harmonize').hide();
        $('#undoHelp').hide();
        $('#tableHeaderHelp').hide();

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
        $('#harmonize').hide();
        $('#undoHelp').hide();
        $('#tableHeaderHelp').hide();

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
        $('#harmonize').hide();
        $('#undoHelp').hide();
        $('#tableHeaderHelp').hide();

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
        $('#harmonize').hide();
        $('#undoHelp').hide();
        $('#tableHeaderHelp').hide();

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


    function harmonizeColumn() {

        $('#firstpage').hide();
        $('#addColumn').hide();
        $('#splitValues').hide();
        $('#extractValues').hide();
        $('#trimSpaces').hide();
        $('#detectDuplicates').hide();
        $('#detectOutlier').hide();
        $('#validationRule').hide();
        $('#harmonize').show();
        $('#undoHelp').hide();
        $('#tableHeaderHelp').hide();

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


    function undoHelp() {

        $('#firstpage').hide();
        $('#addColumn').hide();
        $('#splitValues').hide();
        $('#extractValues').hide();
        $('#trimSpaces').hide();
        $('#detectDuplicates').hide();
        $('#detectOutlier').hide();
        $('#validationRule').hide();
        $('#harmonize').hide();
        $('#undoHelp').show();
        $('#tableHeaderHelp').hide();

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



    function tableHeaderHelp() {

        $('#firstpage').hide();
        $('#addColumn').hide();
        $('#splitValues').hide();
        $('#extractValues').hide();
        $('#trimSpaces').hide();
        $('#detectDuplicates').hide();
        $('#detectOutlier').hide();
        $('#validationRule').hide();
        $('#harmonize').hide();
        $('#undoHelp').hide();
        $('#tableHeaderHelp').show();

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