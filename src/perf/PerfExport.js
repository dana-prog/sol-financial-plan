/**
 * Perf audit exporter — dumps every sheet's formulas, named ranges, and
 * dimensions to a JSON file in Drive so the result can be analyzed offline.
 *
 * Run via: SOL menu → "Export Perf Audit (Formulas + Named Ranges)".
 * After it finishes, an alert shows the Drive file URL. Share that file
 * (or its ID) to run the analysis.
 */

// noinspection JSUnusedGlobalSymbols
function exportPerfAudit() {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheets = spreadsheet.getSheets();

  const dump = {
    spreadsheetId: spreadsheet.getId(),
    spreadsheetName: spreadsheet.getName(),
    exportedAt: new Date().toISOString(),
    locale: spreadsheet.getSpreadsheetLocale(),
    timeZone: spreadsheet.getSpreadsheetTimeZone(),
    recalculationInterval: String(spreadsheet.getRecalculationInterval()),
    namedRanges: _dumpNamedRanges(spreadsheet),
    namedFunctions: _dumpNamedFunctions(spreadsheet.getId()),
    sheets: sheets.map(_dumpSheet),
  };

  const fileName = `perf-audit-${spreadsheet.getName()}-${new Date().getTime()}.json`;
  const file = DriveApp.createFile(fileName, JSON.stringify(dump, null, 2), MimeType.PLAIN_TEXT);

  SpreadsheetApp.getUi().alert(
    'Perf Audit Exported',
    `File: ${fileName}\nID: ${file.getId()}\nURL: ${file.getUrl()}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function _dumpNamedRanges(spreadsheet) {
  return spreadsheet.getNamedRanges().map((namedRange) => {
    const range = namedRange.getRange();
    return {
      name: namedRange.getName(),
      sheet: range.getSheet().getName(),
      a1: range.getA1Notation(),
      numRows: range.getNumRows(),
      numColumns: range.getNumColumns(),
    };
  });
}

function _dumpSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  const result = {
    name: sheet.getName(),
    sheetId: sheet.getSheetId(),
    hidden: sheet.isSheetHidden(),
    maxRows: sheet.getMaxRows(),
    maxColumns: sheet.getMaxColumns(),
    lastRow: lastRow,
    lastColumn: lastColumn,
    frozenRows: sheet.getFrozenRows(),
    frozenColumns: sheet.getFrozenColumns(),
    formulaCount: 0,
    formulas: [],
    conditionalFormatRuleCount: sheet.getConditionalFormatRules().length,
  };

  if (lastRow === 0 || lastColumn === 0) {
    return result;
  }

  const range = sheet.getRange(1, 1, lastRow, lastColumn);
  const formulas = range.getFormulas();

  for (let row = 0; row < formulas.length; row++) {
    const rowFormulas = formulas[row];
    for (let col = 0; col < rowFormulas.length; col++) {
      const formula = rowFormulas[col];
      if (formula) {
        result.formulas.push({
          a1: _a1(row + 1, col + 1),
          f: formula,
        });
        result.formulaCount++;
      }
    }
  }

  return result;
}

/**
 * Named Functions are not exposed by SpreadsheetApp or the Sheets REST API.
 * The xlsx export endpoint preserves them as <definedName> elements inside
 * xl/workbook.xml (LAMBDA-serialized). We fetch that export and parse it.
 * Source: https://gist.github.com/tanaikech/9a9e571ed662e35eec0aa747bb4e025a
 */
function _dumpNamedFunctions(spreadsheetId) {
  try {
    const url = `https://docs.google.com/spreadsheets/export?exportFormat=xlsx&id=${spreadsheetId}`;
    const response = UrlFetchApp.fetch(url, {
      headers: { authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true,
    });
    if (response.getResponseCode() !== 200) {
      return { error: `export http ${response.getResponseCode()}` };
    }

    const blobs = Utilities.unzip(response.getBlob().setContentType(MimeType.ZIP));
    const workbookBlob = blobs.find((blob) => blob.getName() === 'xl/workbook.xml');
    if (!workbookBlob) {
      return { error: 'xl/workbook.xml not found in export' };
    }

    const root = XmlService.parse(workbookBlob.getDataAsString()).getRootElement();
    const definedNamesElement = root.getChild('definedNames', root.getNamespace());
    if (!definedNamesElement) {
      return [];
    }

    return definedNamesElement.getChildren().map((element) => ({
      name: element.getAttribute('name').getValue(),
      definition: element.getValue(),
    }));
  } catch (error) {
    return { error: String(error) };
  }
}

function _a1(row, column) {
  let label = '';
  let rem = column;
  while (rem > 0) {
    const mod = (rem - 1) % 26;
    label = String.fromCharCode(65 + mod) + label;
    rem = Math.floor((rem - 1) / 26);
  }
  return label + row;
}
