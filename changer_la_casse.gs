// Ce script permet de changer la casse des lettres (MAJUSCULES/minuscules/Majuscule Au-Début) dans une feuille de calcul Google.

function onOpen() {
    SpreadsheetApp
        .getUi()
        .createAddonMenu()
        .addItem('MAJUSCULES', 'upper')
        .addItem('minuscules', 'lower')
        .addItem('Majuscule Au-Début', 'proper')
        .addToUi();
}

function lower() {
    run(toLowerCase)
}

function upper() {
    run(toUpperCase)
}

function proper() {
    run(toTitleCase)
}

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

function keepFormulas(sheet, range, formulas) {

    var startRow, startColumn, ui, response;

    startRow = range.getRow();
    startColumn = range.getColumn();

    if (hasFormulas(formulas)) {

        ui = SpreadsheetApp.getUi();
        response = ui.alert('FORMULES TROUVÉS', 'Garder les formules ?', ui.ButtonSet.YES_NO);

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

function hasFormulas(formulas) {
    return formulas.reduce(function (a, b) {
        return a.concat(b);
    })
        .filter(String)
        .length > 0
}
