function SortOut() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("Summary");

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  const table = sheet.getRange(1, 2, lastRow, lastCol - 1).getValues();

  const header = table[0];
  const body = table.slice(1);

  const colIdx = header.map((_, i) => i);

  colIdx.sort((a, b) => {
    const dateA = new Date(header[a]);
    const dateB = new Date(header[b]);
    return dateB - dateA;
  });

  const sortedValues = [];
  sortedValues.push(colIdx.map((i) => header[i]));

  body.forEach((row) => {
    sortedValues.push(colIdx.map((i) => row[i]));
  });

  sheet.getRange(1, 2, lastRow, lastCol - 1).setValues(sortedValues);
}
