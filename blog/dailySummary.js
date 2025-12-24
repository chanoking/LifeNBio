function dailySummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = ss.getSheetByName("Summary");
  let firstRow = 5;

  const sheetNames = [
    "185커큐민",
    "조인트리션",
    "블러드플로우케어",
    "파미로겐",
    "맨드로포즈",
    "헤모웰당",
    "요레스",
    "요로굿",
    "인-칼슘앱솔브",
    "오큘라레이드",
    "위이지케어",
    "비타민D3",
    "리버티엑스",
    "투데이D3",
    "지니어스뉴",
    "데이프로바",
    "이트뮨",
    "그로우뉴",
    "흑본전탕",
  ];

  let lastCol = summary.getLastColumn();

  let totalKeywords = 0,
    totalViews = 0;

  sheetNames.forEach((name) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    const gValues = sheet.getRange(1, 7, 2, 1).getValues().flat(); // G1:G2
    const lastColValues = sheet
      .getRange(1, sheet.getLastColumn(), 2, 1)
      .getValues()
      .flat();

    const [keyword, view, sKey, sView] = [...gValues, ...lastColValues];

    summary
      .getRange(firstRow, lastCol + 1, 4, 1)
      .setValues([[keyword], [view], [sKey], [sView]]);

    (totalKeywords += sKey), (totalViews += sView);

    firstRow += 5;
  });

  const rawDate = new Date();
  const formatted = Utilities.formatDate(
    rawDate,
    Session.getScriptTimeZone(),
    "MM/dd"
  );

  summary
    .getRange(1, lastCol + 1, 3, 1)
    .setValues([[formatted], [totalKeywords], [totalViews]]);

  const lastRow = summary.getLastRow();

  summary
    .getRange(1, lastCol, lastRow, 1)
    .copyTo(summary.getRange(1, lastCol + 1, lastRow, 1), { formatOnly: true });
}
