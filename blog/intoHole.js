function intoHole() {
  const sh = SpreadsheetApp.getActiveSpreadsheet();
  const result = sh.getSheetByName("Hole");
  const sheets = [
    "황둥강둥",
    "오리진",
    "데일리",
    "파이토뉴트리(공식)",
    "푸드케어",
    "허브연구",
  ];

  // ───────────────────────────────
  // 1. Prepare date values
  // ───────────────────────────────
  const tz = Session.getScriptTimeZone();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const weekday = parseInt(Utilities.formatDate(today, tz, "u")); // Mon=1

  const targetDate = new Date(today);
  targetDate.setDate(today.getDate() - (weekday === 1 ? 4 : 1));
  targetDate.setHours(0, 0, 0, 0);

  let lastRow = result.getLastRow();

  let rowsToWrite = [];
  let extraData = [];

  if (lastRow > 1)
    result.getRange(2, 1, lastRow - 1, result.getLastColumn()).clearContent();

  lastRow = result.getLastRow();
  // ───────────────────────────────
  // 2. Process each sheet
  // ───────────────────────────────
  for (const sheetName of sheets) {
    const curSheet = sh.getSheetByName(sheetName);
    if (!curSheet) continue;

    // find last row in column N
    const colN = curSheet
      .getRange(9, 14, curSheet.getLastRow() - 8, 1)
      .getValues();
    let curLastRow = 9;
    for (let i = colN.length - 1; i >= 0; i--) {
      if (colN[i][0] !== "") {
        curLastRow = 9 + i;
        break;
      }
    }

    const dates = curSheet.getRange(9, 14, curLastRow - 8, 1).getValues();

    for (let i = curLastRow; i >= 9; i--) {
      const idx = i - 9;
      const d = dates[idx][0];

      if (!(d instanceof Date)) continue;
      if (d < targetDate) break;
      if (d >= targetDate && d < today) {
        const mainData = curSheet.getRange(i, 14, 1, 5).getValues()[0];
        const fileName = curSheet.getRange(i, 13).getValue();
        const keyword = curSheet.getRange(i, 5).getValue();
        const type = curSheet.getRange(i, 4).getValue();

        const split = fileName ? fileName.split("_") : [];

        rowsToWrite.push(mainData);
        extraData.push([split, keyword, type]);
      }
    }
  }

  rowsToWrite.reverse();
  extraData.reverse();

  // ───────────────────────────────
  // 3. Write collected rows in one batch
  // ───────────────────────────────

  for (let i = 0; i < rowsToWrite.length; i++) {
    result.getRange(lastRow + 1, 8, 1, 5).setValues([rowsToWrite[i]]);

    const split = extraData[i][0];
    const keyword = extraData[i][1];
    const type = extraData[i][2];

    if (split.length > 0) {
      result.getRange(lastRow + 1, 15, 1, split.length).setValues([split]);
      let str = "";
      str = `${split[0]}/${split[4]}_${split[5]}/${split[2]}_${split[3]}${split[7]}_${split[8]}_${split[9]}_${split[13]}/${split[11]}_${rowsToWrite[i][3]}`;
      result.getRange(lastRow + 1, 13, 1, 1).setValue(str);
    }

    result.getRange(lastRow + 1, 26).setValue(keyword);
    result.getRange(lastRow + 1, 29).setValue(type);

    lastRow++;
  }

  // ───────────────────────────────
  // 4. Fix URLs (batch replace)
  // ───────────────────────────────
  const map = {
    https: "http",
    "m.": "",
  };

  let urls = result.getRange(2, 11, result.getLastRow() - 1, 1).getValues();
  let types = result.getRange(2, 29, result.getLastRow() - 1, 1).getValues();

  urls = urls.map((row) => {
    let url = row[0];
    if (typeof url === "string") {
      url = url.replace(/https|m\./g, (m) => map[m]);
    }
    return [url];
  });

  types = types.map((t) => {
    let type = t[0].replace("스블", "서브");

    return [type];
  });

  result.getRange(2, 11, urls.length, 1).setValues(urls);
  result.getRange(2, 27, types.length, 1).setValues(types);

  let dates = result.getRange(2, 8, result.getLastRow() - 1, 1).getValues();
  let items = result.getRange(2, 20, result.getLastRow() - 1, 1).getValues();
  let comb = result.getRange(2, 23, result.getLastRow() - 1, 4).getValues();

  result.getRange("A:A").setNumberFormat("MM/dd");

  result.getRange(2, 1, types.length, 1).setValues(dates);
  result.getRange(2, 14, types.length, 1).setValues(dates);
  result.getRange(2, 2, types.length, 1).setValue("자사최블");
  result.getRange(2, 3, types.length, 1).setValues(items);
  result.getRange(2, 4, types.length, 4).setValues(comb);
  result
    .getRange(2, result.getLastColumn() - 2, types.length, 1)
    .setValue("자사최블");

  const urlTitle = result
    .getRange(2, 11, result.getLastRow() - 1, 2)
    .getValues();
  result
    .getRange(2, result.getLastColumn() - 1, result.getLastRow() - 1, 2)
    .setValues(urlTitle);

  result
    .getRange(2, 1, result.getLastRow() - 1, result.getLastColumn())
    .sort({ column: 1, ascending: true });
}
