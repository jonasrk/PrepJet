/**
 * Created by jonas on 6/23/16.
 */

// Pass the checkbox name to the function
function getCheckedBoxes(chkboxName) {
    var checkboxes = document.getElementsByName(chkboxName);
    var checkboxesChecked = [];
    // loop over them all
    for (var i=0; i<checkboxes.length; i++) {
        // And stick the checked ones onto an array...
        if (checkboxes[i].checked) {
            checkboxesChecked.push(checkboxes[i]);
        }
    }
    // Return the array if it is non-empty, or null
    return checkboxesChecked.length > 0 ? checkboxesChecked : null;
}


// create a ms-ChoiceField html element (input + div + label + span) for every column in the selected table
function addNewCheckboxToContainer (id, name, container) {

    var el =  document.createElement("input");
    el.id = id;
    el.name = name;
    el.className = "ms-ChoiceField-input";
    el.setAttribute("type", "checkbox");

    var div = document.createElement("div");

    var label =  document.createElement("label");
    label.className = "ms-ChoiceField-field";
    label.setAttribute("for", id);

    var span =  document.createElement("span");
    span.className = "ms-Label";
    span.textContent = id;

    label.appendChild(span);
    div.appendChild(el);
    div.appendChild(label);

    document.getElementById(container).appendChild(div);

}


function getCharFromNumber (number) {

    if (number == 1) {
        return 'A';
    } else if (number == 2) {
        return 'B';
    } else if (number == 3) {
        return 'C';
    } else if (number == 4) {
        return 'D';
    } else if (number == 5) {
        return 'E';
    } else if (number == 6) {
        return 'F';
    } else if (number == 7) {
        return 'G';
    } else if (number == 8) {
        return 'H';
    } else if (number == 9) {
        return 'I';
    } else if (number == 10) {
        return 'J';
    } else if (number == 11) {
        return 'K';
    } else if (number == 12) {
        return 'L';
    } else if (number == 13) {
        return 'M';
    } else if (number == 14) {
        return 'N';
    } else if (number == 15) {
        return 'O';
    } else if (number == 16) {
        return 'P';
    } else if (number == 17) {
        return 'Q';
    } else if (number == 18) {
        return 'R';
    } else if (number == 19) {
        return 'S';
    } else if (number == 20) {
        return 'T';
    } else if (number == 21) {
        return 'U';
    } else if (number == 22) {
        return 'V';
    } else if (number == 23) {
        return 'W';
    } else if (number == 24) {
        return 'X';
    } else if (number == 25) {
        return 'Y';
    } else if (number == 26) {
        return 'Z';
    }

}


// Helper function to add and format content in the workbook
function addContentToWorksheet(sheetObject, rangeAddress, displayText) {
    var range = sheetObject.getRange(rangeAddress);
    range.values = displayText;
    range.merge();
}