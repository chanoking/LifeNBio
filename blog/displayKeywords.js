function DisplayKeywords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetById(1695242273);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  const items = sheet.getRange(3, 1, 1, lastCol).getValues()[0];
  const name = sheet.getName(); // FIXED

  // Clear old contents
  if (lastRow > 1) {
    sheet.getRange(4, 1, lastRow - 1, lastCol).clearContent();
  }

  for (let i = 0; i < items.length; i++) {
    const sheetName = items[i];
    const curSheet = ss.getSheetByName(sheetName);
    if (!curSheet) continue;

    const curLastRow = curSheet.getLastRow();
    const curLastCol = curSheet.getLastColumn();
    if (curLastRow < 4) continue;

    const keys = curSheet.getRange("A4:A" + curLastRow).getValues();
    const v = curSheet.getRange("B4:B" + curLastRow).getValues();
    const values = curSheet
      .getRange(4, curLastCol, curLastRow - 3, 1)
      .getValues();

    let keywords;
    let views;
    if (name === "노출") {
      keywords = keys.filter((_, idx) => values[idx][0] === 1);
      views = v.filter((_, idx) => values[idx][0] === 1);
    } else {
      keywords = keys.filter((_, idx) => values[idx][0] === 0);
      views = v.filter((_, idx) => values[idx][0] === 0);
    }

    let sumView1 = views.reduce((acc, cur) => {
      if (cur[0] === "" || cur[0] === null) {
        return (acc += 0);
      }
      return acc + cur[0];
    }, 0);

    // ❗ Avoid writing empty range
    if (keywords.length > 0) {
      sheet.getRange(4, i + 1, keywords.length, 1).setValues(keywords);
    }

    sheet.getRange(1, i + 1, 1, 1).setValue(keywords.length);
    sheet.getRange(2, i + 1, 1, 1).setValue(sumView1);
  }
}
