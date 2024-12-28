// This script is used to change the case of letters (UPPERCASE/lowercase/Title-Case) in a Google spreadsheet.

// onOpen is used to create a menu in the Google spreadsheet
function onOpen() {
    SpreadsheetApp
        .getUi()
        .createAddonMenu()
        .addItem('UPPERCASE', 'upper')
        .addItem('lowercase', 'lower')
        .addItem('Title-Case', 'proper')
        .addToUi();
}

// function called when the user selects the UPPERCASE option in the menu
function upper() {
    run(toUpperCase)
}

// function called when the user selects the lowercase option in the menu
function lower() {
    run(toLowerCase)
}
// function called when the user selects the Title-Case option in the menu
function proper() {
    run(toTitleCase)
}

// run is used to apply the function fn (toUpperCase, toLowerCase, toTitleCase) 
// to the selected range.
function run(fn) {
    var r, s, v, f;

    s = SpreadsheetApp.getActiveSheet(),
        r = s.getActiveRange()
    v = r.getValues();
    f = r.getFormulas()

    r.setValues(
        v.map(function (ro) {
            return ro.map(function (el) {
                return !el ? null : typeof el !== 'string' && el ? el : fn(el);
            })
        })
    )
    keepFormulas(s, r, f);
}

// the actual functions that change the case of the letters 
function toUpperCase(str) {
    return str.toUpperCase();
}

function toLowerCase(str) {
    return str.toLowerCase();
}

function toTitleCase(str) {
    return str.toLowerCase().replace(/(?:^|[\s-/])\w/g,
        function (match) {
            return match.toUpperCase();
        });
}

// keepFormulas is used to ask the user if they want to keep 
// the formulas (if any) in the selected range.
function keepFormulas(sheet, range, formulas) {

    var startRow, startColumn, ui, response;

    startRow = range.getRow();
    startColumn = range.getColumn();

    if (hasFormulas(formulas)) {

        ui = SpreadsheetApp.getUi();
        response = ui.alert('FORMULAS FOUND', 'Keep the formulas?', ui.ButtonSet.YES_NO);

        if (response == ui.Button.YES) {
            formulas.forEach(function (r, i) {
                r.forEach(function (c, j) {
                    if (c) sheet.getRange((startRow + i), (startColumn + j))
                        .setFormula(formulas[i][j])
                })
            })
        }
    }
}

// hasFormulas is used to check if the selected range contains formulas
function hasFormulas(formulas) {
    return formulas.reduce(function (a, b) {
        return a.concat(b);
    })
        .filter(String)
        .length > 0
}
