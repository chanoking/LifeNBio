function smallChange() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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

  const d = new Date();
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() - 1);

  sheetNames.forEach((name) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    const lastCol = sheet.getLastColumn();
    sheet.getRange(3, lastCol).setValue(d);
  });
}
