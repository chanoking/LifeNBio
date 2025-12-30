function sortAndMakeLines() {
  const ss = SpreadsheetApp.getActive().getSheetByName("키챌키워드");

  const lastRow = ss.getLastRow();
  const lastCol = ss.getLastColumn();

  for (let r = 12; r <= lastRow; r++) {
    const msA = ss.getRange(r + 1, 4, 1, 1).getValue();
    const msC = ss.getRange(r, 4, 1, 1).getValue();

    if ((msC === "메인" && msA === msC) || (msC === "서브" && msA !== msC)) {
      ss.getRange(r, 1, 1, lastCol).setBorder(
        null,
        null,
        true,
        null,
        null,
        null
      );
    } else {
    }
  }
}
