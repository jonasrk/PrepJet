(function(){
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function(reason){
        jQuery(document).ready(function(){
            app.initialize();

            jQuery('#test').click(getDataFromSelection);
            //jQuery('#replace-checked-values').click(replaceCheckedValues);

            }
        });
    };

    // Reads data from current document selection and displays a notification
    function getDataFromSelection(){
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function(result){

                var txt = document.createElement("label");
                txt.innerHTML = "testtestetst";
                document.getElementById('explanation').appendChild(txt);
                console.log("print test");

            }
        );
    }

    function replaceCheckedValues(){

        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function(result){

            }
        );

    }

})();
