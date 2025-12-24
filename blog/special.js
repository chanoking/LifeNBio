function special() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("통합");
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

  const lastRow = rawSheet.getLastRow();

  // master keywords → uppercase, remove blanks
  const keywordsUpper = rawSheet
    .getRange(3, 2, lastRow - 2, 1)
    .getValues()
    .flat()
    .filter((k) => k && k.toString().trim() !== "")
    .map((k) => k.toString().toUpperCase());

  const unusedKeywords = [];

  sheetNames.forEach((name) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    const curLastRow = sheet.getLastRow();

    // product sheet keywords → uppercase, remove blanks
    const curKeywords = sheet
      .getRange(4, 1, curLastRow - 3, 1)
      .getValues()
      .flat()
      .filter((k) => k && k.toString().trim() !== "")
      .map((key) => key.toString().toUpperCase());

    curKeywords.forEach((cK) => {
      if (!keywordsUpper.includes(cK)) {
        unusedKeywords.push([cK]);
      }
    });
  });

  if (unusedKeywords.length > 0) {
    rawSheet.getRange(3, 9, unusedKeywords.length, 1).setValues(unusedKeywords);
  }
}
