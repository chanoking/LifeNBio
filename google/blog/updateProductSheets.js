function updateProductSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const original = ss.getSheetByName("통합");
  const priority = ss.getSheetByName("우선순위");
  if (!original) return;

  // ─────────────────────────────────────
  // 1) Read original data (batch)
  // ─────────────────────────────────────
  const lastRow = original.getLastRow();
  if (lastRow < 3) return;

  const productsRaw = original.getRange(`A3:A${lastRow}`).getValues();
  const views = original.getRange(`E3:E${lastRow}`).getValues();
  const visibles = original.getRange(`F3:F${lastRow}`).getValues();
  const blocks = original.getRange(`G3:G${lastRow}`).getValues();
  const keywordsRaw = original.getRange(`B3:B${lastRow}`).getValues();

  const rawProductsArr = [
    "조인트리션",
    "블러드",
    "비타민",
    "칼슘앱솔브",
    "이트뮨",
    "지니어스뉴",
    "투데이",
  ];
  const refineProductsArr = [
    "조인트리션",
    "블러드플로우케어",
    "비타민D3",
    "인-칼슘앱솔브",
    "이트뮨",
    "지니어스뉴",
    "투데이D3",
  ];
  const products = productsRaw.map((row) => {
    let refineP = row[0] || "";
    rawProductsArr.some((p, i) => {
      if (refineP.includes(p)) {
        refineP = refineProductsArr[i];
        return true; // stop loop
      }
      return false;
    });
    return [refineP];
  });

  const keywords = keywordsRaw.map((row) => {
    const k = (row[0] || "").toString();
    return [k.toUpperCase()];
  });

  original.getRange(`B3:B${lastRow}`).setValues(keywords);

  // Priority sheet values
  const pLastRow = priority ? priority.getLastRow() : 1;
  const priorityKeywords =
    priority && pLastRow >= 2
      ? priority.getRange(`A2:A${pLastRow}`).getValues()
      : [];
  const priorities =
    priority && pLastRow >= 2
      ? priority.getRange(`B2:B${pLastRow}`).getValues()
      : [];

  // ─────────────────────────────────────
  // 2) Build lookup maps (keyword -> view / rank / block / priority)
  // Use uppercase keys consistently
  // ─────────────────────────────────────
  const viewMap = Object.create(null);
  const rankMap = Object.create(null);
  const blockMap = Object.create(null);

  for (let i = 0; i < keywords.length; i++) {
    const key = (keywords[i][0] || "").toString();
    if (!key) continue;
    if (!(key in viewMap)) {
      // keep first encountered
      viewMap[key] = Number(views[i][0]) || 0;
      rankMap[key] = Number(visibles[i][0]) || 0;
      blockMap[key] = Number(blocks[i][0]) || 0;
    }
  }

  const priorityMap = Object.create(null);
  for (let i = 0; i < priorityKeywords.length; i++) {
    const key = (priorityKeywords[i][0] || "").toString();
    if (!key) continue;
    if (!(key in priorityMap)) priorityMap[key] = priorities[i][0];
  }

  // ─────────────────────────────────────
  // 3) Product sheet list
  // ─────────────────────────────────────
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

  // ─────────────────────────────────────
  // 4) Iterate product sheets (optimized)
  // ─────────────────────────────────────
  sheetNames.forEach((name) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    let lastRowSheet = sheet.getLastRow();
    if (lastRowSheet < 4) return; // nothing to update

    sheet
      .getRange(4, 1, lastRowSheet - 3, 1)
      .setBackground(null)
      .setFontWeight("normal");

    // Read existing keywords (A4:A...)
    let curKeywords = sheet.getRange(4, 1, lastRowSheet - 3, 1).getValues(); // [[k],...]

    // Normalize current keywords to uppercase for comparison
    const curKeywordsUpper = curKeywords.map((r) =>
      (r[0] || "").toString().toUpperCase()
    );

    // Append missing keywords from original (products -> keywords)
    const rowsToAppend = [];
    for (let i = 0; i < products.length; i++) {
      if ((products[i][0] || "") === name) {
        // exact match to sheet name
        const origKey = (keywords[i][0] || "").toString();
        const up = origKey.toUpperCase();
        if (up && !curKeywordsUpper.includes(up)) {
          rowsToAppend.push([origKey]); // preserve original case for storing
          curKeywordsUpper.push(up); // prevent duplicates within this run
        }
      }
    }

    if (rowsToAppend.length > 0) {
      sheet
        .getRange(lastRowSheet + 1, 1, rowsToAppend.length, 1)
        .setValues(rowsToAppend);
      lastRowSheet = sheet.getLastRow();
      curKeywords = sheet.getRange(4, 1, lastRowSheet - 3, 1).getValues();
    }

    // Remove duplicate keywords in sheet (keep first instance)
    const seen = Object.create(null);
    const deleteRows = []; // collect 1-based row numbers to delete (sheet row)
    for (let i = 0; i < curKeywords.length; i++) {
      const raw = (curKeywords[i][0] || "").toString();
      const up = raw.toUpperCase();
      const sheetRow = 4 + i;
      if (!up) {
        // empty -> mark to delete (optional). If you prefer to keep, skip this block.
        // deleteRows.push(sheetRow);
        continue;
      }
      if (!seen[up]) {
        seen[up] = true;
      } else {
        deleteRows.push(sheetRow);
      }
    }
    // delete from bottom to top to preserve indices
    if (deleteRows.length) {
      deleteRows.sort((a, b) => b - a).forEach((rnum) => sheet.deleteRow(rnum));
      lastRowSheet = sheet.getLastRow();
      curKeywords = sheet.getRange(4, 1, lastRowSheet - 3, 1).getValues();
    }

    // Recompute number of keyword rows after potential modifications
    const nRows = curKeywords.length;
    if (nRows === 0) return;

    // Ensure there's an empty column at the end to write new rank column (if last column not empty)
    let lastCol = sheet.getLastColumn();
    // check the top 3 rows in that lastCol - if any has content, add a new column
    const topCells = sheet.getRange(1, lastCol, 3, 1).getValues().flat();
    const topUsed = topCells.some((v) => v !== "" && v !== null);
    if (topUsed) {
      sheet.insertColumnAfter(lastCol);
      lastCol++;
    }

    // Prepare result arrays
    const viewResult = new Array(nRows);
    const blockResult = new Array(nRows);
    const solitaryResult = new Array(nRows);
    const rankResult = new Array(nRows);
    const priorityResult = new Array(nRows);

    let sumViews = 0;
    let sumKeywords = 0;
    let totalViews = 0;
    let totalKeywords = 0;

    for (let i = 0; i < nRows; i++) {
      const keyRaw = (curKeywords[i][0] || "").toString();
      const key = keyRaw.toUpperCase();

      totalKeywords++;
      totalViews += Number(viewMap[key]) || 0;

      viewResult[i] = [viewMap[key] || ""];
      blockResult[i] = [blockMap[key] || ""];
      priorityResult[i] = [priorityMap[key] || ""];

      const isBlocked = Number(blockMap[key]) === 1;
      const isRanked = Number(rankMap[key]) === 1;

      if (isBlocked) {
        solitaryResult[i] = [0];
      } else if (isRanked) {
        solitaryResult[i] = [1];
      } else {
        solitaryResult[i] = [""];
      }

      rankResult[i] = [rankMap[key] || 0];

      if (isRanked) {
        sumViews += Number(viewMap[key]) || 0;
        sumKeywords++;
      }
    }

    // Top summary: sumKeywords, sumViews, today
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    sheet
      .getRange(1, lastCol, 3, 1)
      .setValues([[sumKeywords], [sumViews], [today]]);

    // Write bulk results to sheet
    const writeRows = lastRowSheet - 3;
    // Defensive: if sizes mismatch, adjust to lengths
    const rowsToWriteCount = Math.min(writeRows, viewResult.length);

    sheet
      .getRange(4, 2, rowsToWriteCount, 1)
      .setValues(viewResult.slice(0, rowsToWriteCount));
    sheet
      .getRange(4, 4, rowsToWriteCount, 1)
      .setValues(blockResult.slice(0, rowsToWriteCount));
    sheet
      .getRange(4, 5, rowsToWriteCount, 1)
      .setValues(solitaryResult.slice(0, rowsToWriteCount));
    sheet
      .getRange(4, lastCol, rowsToWriteCount, 1)
      .setValues(rankResult.slice(0, rowsToWriteCount));
    sheet
      .getRange(4, 3, rowsToWriteCount, 1)
      .setValues(priorityResult.slice(0, rowsToWriteCount));

    // Also set overall totals at fixed position (example used earlier: (1,7) two rows)
    sheet.getRange(1, 7, 2, 1).setValues([[totalKeywords], [totalViews]]);
    const aVals = sheet.getRange(4, lastCol, rowsToWriteCount, 1).getValues();
    const bVals = sheet
      .getRange(4, lastCol - 1, rowsToWriteCount, 1)
      .getValues();

    // helper function
    const isNumber = (v) => typeof v === "number" && !isNaN(v);

    for (let i = 0; i < rowsToWriteCount; i++) {
      const a = aVals[i][0]; // new value
      const b = bVals[i][0]; // old value

      const row = i + 4;

      // --- Case 1: identical → keep default color
      if (a === b) continue;

      // --- Case 2: previous was blank, now valid number → NEW item
      if ((b === "" || b === null) && isNumber(a)) {
        sheet.getRange(row, 1).setBackground("#d9ead3").setFontWeight("bold");
        continue;
      }

      // --- Case 3: value increased
      if (isNumber(a) && isNumber(b) && a > b) {
        sheet.getRange(row, 1).setBackground("#D9EAF7").setFontWeight("bold"); // light blue
        continue;
      }

      // --- Case 4: value decreased
      if (isNumber(a) && isNumber(b) && a < b) {
        sheet.getRange(row, 1).setBackground("#F8D7DA").setFontWeight("bold");
        continue;
      }

      // --- Otherwise: nothing special
    }
  });
}
