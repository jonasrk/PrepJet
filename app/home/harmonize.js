function resultOK() {
    document.getElementById('resultDialog').style.visibility = 'hidden';
    window.location = "harmonize.html";
}

function resultClose() {
    document.getElementById('resultDialog').style.visibility = 'hidden';
    window.location = "harmonize.html";
}

function redirectHome() {
    window.location = "mac_start.html";
}

(function () {
    // 'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        jQuery(document).ready(function () {

            app.initialize();

            $('.harmonize_drop').Dropdown();

            jQuery('#harmonize').click(harmonize);
            jQuery('#homeButton').click(redirectHome);
            jQuery('#resultOk').click(resultOK);
            jQuery('#resultOk').click(resultClose);


        });
    };

    function getSelectedData(callback) {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Matrix, { valueFormat: Office.ValueFormat.Formatted },
        function (result) {
            if (result.status == "succeeded") {
                callback(result.value);
            }
            else {
                console.log("error");
            }
        });
    }


    function getHarmItem(item) {
        var harmo = document.getElementById('harmonize_options').value;
        if (harmo == "allupper") {
            var harm_string = item.toUpperCase();
        }
        if (harmo == "alllower") {
            var harm_string = item.toLowerCase();
        }
        if (harmo == "firstupper") {
            /*var tmp = item.toLowerCase().split(" ");
            var tmp_upper = [];
            for (var runtmp = 0; runtmp < tmp.length; runtmp++) {
                tmp_upper.push(tmp[runtmp].charAt(0).toUpperCase() + tmp[runtmp].slice(1));
            }
            var harm_string = tmp_upper[0];
            for (var runtmp = 1; runtmp < tmp_upper.length; runtmp++) {
                harm_string = harm_string.concat(" ", tmp_upper[runtmp]);
            }
            harm_string = [harm_string];*/
            var tmp = item.toLowerCase();
            var harm_string = tmp;
        }
        if (harmo == "oneupper") {
            var tmp = item.split(" ");
            var tmp_upper = [];
            tmp_upper.push(tmp[0].charAt(0).toUpperCase() + tmp[0].slice(1).toLowerCase());
            for (var runtmp = 1; runtmp < tmp.length; runtmp++) {
                tmp_upper.push(tmp[runtmp].charAt(0) + tmp[runtmp].slice(1).toLowerCase());
            }

            var harm_string = tmp_upper[0];
            for (var runtmp = 1; runtmp < tmp_upper.length; runtmp++) {
                harm_string = harm_string.concat(" ", tmp_upper[runtmp]);
            }
            harm_string = [harm_string];
        }

        return harm_string;
    }

    function harmonize() {

        getSelectedData(function(result){

                if (result != null) {
                    var countStr = 0;
                    var harm_array = result.map(function (item) {
                        return item.map(function (item) {
                            if (item) {
                                var newitem = getHarmItem(item);
                                if (item != newitem) {
                                    countStr++;
                                }
                                return newitem;
                            }
                        });
                    });
                }
                Office.context.document.setSelectedDataAsync(harm_array, { valueFormat: Office.ValueFormat.Formatted }, function(result){
                        if (result.status == "succeeded") {
                            var txt = document.createElement("p");
                            txt.className = "ms-font-xs ms-embedded-dialog__content__text";
                            txt.innerHTML = "PrepJet successfully changed the cases of  " + countStr + " data entries.";
                            document.getElementById('resultText').appendChild(txt);
                            document.getElementById('resultDialog').style.visibility = 'visible';
                        } else {
                            console.log("An error occured. Please select a range and try again.");
                        }
                    });
        });
    }


})();