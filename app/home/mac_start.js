(function () {
    // 'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            app.initialize();

            $zopim(function() {
                $zopim.livechat.window.hide();
            });

            /*$('#splitValHome').click(home);
            $('#addColHome').click(home);
            $('#extractValHome').click(home);
            $('#trimSpacesHome').click(home);
            $('#detDupHome').click(home);
            $('#detOutHome').click(home);
            $('#valRuleHome').click(home);
            $('#harmonizeHome').click(home);
            $('#undoHelpHome').click(home);
            $('#tableHeaderHome').click(home);
            $('#customIncon').click(home);*/

        });
    };

})();