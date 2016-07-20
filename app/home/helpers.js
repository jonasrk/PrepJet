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

// Pass the checkbox name to the function
function getAllCheckBoxes(chkboxName) {
    var checkboxes = document.getElementsByName(chkboxName);
    var checkboxesChecked = [];
    // loop over them all
    for (var i=0; i<checkboxes.length; i++) {
        checkboxesChecked.push(checkboxes[i]);
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


function addDropdown (k) {
    var div = document.createElement("div");
    div.id = "condition" + k;
    document.getElementById("condition_holder").appendChild(div);

    var div_drop = document.createElement("div");
    div_drop.className = "ms-Dropdown table_simple" + k;
    div_drop.id = "simple_dropdown" + k;
    div.appendChild(div_drop);

    var lab = document.createElement('label');
    lab.className = "ms-Label";
    lab.innerHTML = "Select column";
    div_drop.appendChild(lab);

    var elemi = document.createElement("i");
    elemi.className = "ms-Dropdown-caretDown ms-Icon ms-Icon--caretDown";
    div_drop.appendChild(elemi);

    var sel = document.createElement("select");
    sel.id = "column_simple" + k;
    sel.className = "ms-Dropdown-select";
    div_drop.appendChild(sel);
}

function addOperator(k) {
    var div_drop = document.createElement("div");
    div_drop.className = "ms-Dropdown dropdown_table" + k;
    document.getElementById('condition' + k).appendChild(div_drop);

    var lab = document.createElement('label');
    lab.className = "ms-Label";
    lab.innerHTML = "Select operator";
    div_drop.appendChild(lab);

    var elemi = document.createElement("i");
    elemi.className = "ms-Dropdown-caretDown ms-Icon ms-Icon--caretDown";
    div_drop.appendChild(elemi);

    var sel = document.createElement("select");
    sel.id = "if_operator" + k;
    sel.className = "ms-Dropdown-select";
    div_drop.appendChild(sel);

    var el1 = document.createElement("option");
    el1.value = "equal" + k;
    el1.textContent = "=";
    sel.appendChild(el1);

    var el2 = document.createElement("option");
    el2.value = "smaller" + k;
    el2.textContent = "<";
    sel.appendChild(el2);

    var el3 = document.createElement("option");
    el3.value = "greater" + k;
    el3.textContent = ">";
    sel.appendChild(el3);

    var el4 = document.createElement("option");
    el4.value = "inequal" + k;
    el4.textContent = "!=";
    sel.appendChild(el4);

    var el5 = document.createElement("option");
    el5.value = "between" + k;
    el5.textContent = "between";
    sel.appendChild(el5);

    var el6 = document.createElement("option");
    el6.value = "notbetween" + k;
    el6.textContent = "not between";
    sel.appendChild(el6);

    var el7 = document.createElement("option");
    el7.value = "inlist" + k;
    el7.textContent = "in (list)";
    sel.appendChild(el7);
}


function addTextField(k) {
    var div_drop = document.createElement("div");
    div_drop.className = "ms-TextField";
    div_drop.id = "delimiter_beginning" + k;
    document.getElementById('condition' + k).appendChild(div_drop);

    var lab = document.createElement('label');
    lab.className = "ms-Label";
    lab.innerHTML = "Enter condition";
    div_drop.appendChild(lab);

    var input = document.createElement("input");
    input.id = "if_condition" + k;
    input.className = "ms-TextField-field";
    input.type = "text";
    div_drop.appendChild(input);
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


function highlightContentInWorksheet(sheetObject, rangeAddress, color) {
    var range = sheetObject.getRange(rangeAddress);
    range.format.font.color = color;
    range.merge();
}


function highlightCellInWorksheet(sheetObject, rangeAddress, color) {
    var range = sheetObject.getRange(rangeAddress);
    range.format.fill.color = color;
    range.merge();
}


function getRandomColor() {
    var letters = '0123456789ABCDEF'.split('');
    var color = '#';
    for (var i = 0; i < 6; i++ ) {
        color += letters[Math.floor(Math.random() * 16)];
    }
    return color;
}

