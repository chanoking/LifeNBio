function pasteRank() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = ss.getSheetByName("키챌키워드");
  var sourceSheet = ss.getSheetByName("변환용");

  // Get current keywords and types
  var lastRow = currentSheet.getLastRow();
  var currentKeywordsRange = currentSheet.getRange("B12:B" + lastRow);
  var currentSthRange = currentSheet.getRange("H12:H" + lastRow);
  var currentTypeRange = currentSheet.getRange("G12:G" + lastRow);
  var currentKeywords = currentKeywordsRange.getValues();
  var currentTypes = currentTypeRange.getValues();
  var currentSth = currentSthRange.getValues();

  // Get source data
  var sourceData = sourceSheet.getDataRange().getValues();
  var rankMap = {};
  var viewMap = {};
  var dateMap = {};

  for (var i = 0; i < sourceData.length; i++) {
    var key = sourceData[i][0];
    if (key !== "") {
      viewMap[key] = sourceData[i][1];
      rankMap[key] = sourceData[i][2];
      dateMap[key] = sourceData[i][3];
    }
  }

  // Insert new column for rank if needed
  if (currentSheet.getRange("O1").isBlank()) {
    currentSheet.insertColumnBefore(15);
  }

  // Prepare outputs and counts
  var rankOutput = [];
  var viewOutput = [];
  var dateOutput = [];
  let typeOutput = [];
  var countA = 0,
    countB = 0;

  // console.log(currentKeywords.length)
  for (var j = 0; j < currentKeywords.length; j++) {
    if (
      currentTypes[j][0] === "키챌보장" &&
      currentSth[j][0] !== "제품 일시 중단"
    ) {
      countB++;
    } else if (
      currentTypes[j][0] === "키챌건바이" &&
      currentSth[j][0] !== "제품 일시 중단"
    ) {
      countA++;
    }

    var kw = currentKeywords[j][0];
    rankOutput.push([rankMap.hasOwnProperty(kw) ? rankMap[kw] : ""]);
    viewOutput.push([viewMap.hasOwnProperty(kw) ? viewMap[kw] : ""]);
    dateOutput.push([dateMap.hasOwnProperty(kw) ? dateMap[kw] : ""]);
  }

  // Fill counts
  currentSheet.getRange("E3:G3").setValues([[countA + countB, countA, countB]]);

  // Paste ranks and views
  currentSheet.getRange(12, 15, rankOutput.length, 1).setValues(rankOutput);
  currentSheet.getRange(12, 3, viewOutput.length, 1).setValues(viewOutput);
  currentSheet.getRange(12, 10, dateOutput.length, 1).setValues(dateOutput);

  // Set today's date (without time)
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  currentSheet.getRange("O11").setValue(today);

  // Copy formatting from column P (16) to O (15)
  var sourceColumn = 16;
  var targetColumn = 15;
  var sourceRange = currentSheet.getRange(1, sourceColumn, lastRow);
  var targetRange = currentSheet.getRange(1, targetColumn, lastRow);
  sourceRange.copyTo(targetRange, { formatOnly: true });

  // Highlight keywords with rank > 0 (light yellow and bold)
  var colors = [];
  var fonts = [];

  let typesQuotes = currentSheet
    .getRange(12, 4, rankOutput.length, 2)
    .getValues();
  for (var k = 0; k < rankOutput.length; k++) {
    var rank = rankOutput[k][0];
    let type = typesQuotes[k][0];
    let quote = typesQuotes[k][1];
    colors.push([
      rank !== "" && rank === 0 && type === "메인" && typeof quote === "number"
        ? "#F08080"
        : null,
    ]);
    fonts.push([
      rank !== "" && rank === 0 && type === "메인" && typeof quote === "number"
        ? "bold"
        : "normal",
    ]);
  }
  currentSheet.getRange(12, 2, colors.length, 1).setBackgrounds(colors);
  currentSheet.getRange(12, 2, fonts.length, 1).setFontWeights(fonts);

  // Count rank > 0 by type
  var currentRankRange = currentSheet.getRange("O12:O" + lastRow);
  var currentRanks = currentRankRange.getValues();
  var countAA = 0,
    countBB = 0;
  for (var i = 0; i < currentRanks.length; i++) {
    if (
      currentRanks[i][0] > 0 &&
      currentTypes[i][0] === "키챌건바이" &&
      currentSth[i][0] !== "제품 일시 중단"
    ) {
      countAA++;
    }
    if (
      currentRanks[i][0] > 0 &&
      currentTypes[i][0] === "키챌보장" &&
      currentSth[i][0] !== "제품 일시 중단"
    ) {
      countBB++;
    }
  }
  currentSheet
    .getRange("E5:G5")
    .setValues([[countAA + countBB, countAA, countBB]]);

  // Function to convert column number to letter
  function columnToLetter(column) {
    var temp = "";
    var letter = "";
    while (column > 0) {
      temp = (column - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      column = Math.floor((column - 1) / 26);
    }
    return letter;
  }

  var lastColLetter = columnToLetter(currentSheet.getLastColumn());
  var dateColumns = currentSheet
    .getRange("O11:" + lastColLetter + "11")
    .getValues();
  var lastWeekCnt = 0;
  var lastWeek = new Date(today);
  lastWeek.setDate(lastWeek.getDate() - 7);
  lastWeek.setHours(0, 0, 0, 0);
  //console.log(lastWeek)
  var lastWeekCnt = 0;
  for (var c = 0; c < dateColumns[0].length; c++) {
    if (
      dateColumns[0][c] instanceof Date &&
      dateColumns[0][c].getTime() === lastWeek.getTime()
    ) {
      let lastWeekRanks = currentSheet
        .getRange(12, c + 15, currentRanks.length, 1)
        .getValues();

      for (var r = 0; r < lastWeekRanks.length; r++) {
        if (lastWeekRanks[r][0] > 0) lastWeekCnt++;
      }
    }
  }

  // Set last week counts and ratios
  var ratio =
    countAA + countBB + lastWeekCnt !== 0
      ? (countAA + countBB - lastWeekCnt) / (countAA + countBB + lastWeekCnt)
      : 0;
  currentSheet.getRange("H3:I3").setValues([[lastWeekCnt, ratio]]);
  currentSheet
    .getRange("H5:I5")
    .setValues([
      [
        (countAA + countBB) / (countA + countB),
        lastWeekCnt / (countA + countB),
      ],
    ]);

  const sendDates = currentSheet.getRange(12, 9, lastRow - 11, 1).getValues();
  const putoutDates = currentSheet
    .getRange(12, 10, lastRow - 11, 1)
    .getValues();

  currentSheet
    .getRange(12, 9, lastRow - 11, 2)
    .setBackground(null)
    .setFontWeight("normal");

  for (let i = 0; i < sendDates.length; i++) {
    const sendD = sendDates[i][0];
    const putoutD = putoutDates[i][0];
    const row = 12 + i;
    const rank = rankOutput[i][0];

    if (!(sendD instanceof Date)) continue;

    if (putoutD instanceof Date && sendD <= putoutD) {
      currentSheet
        .getRange(row, 10)
        .setBackground("#D9EAF7")
        .setFontWeight("bold");
      continue;
    }

    if (rank > 0) {
      currentSheet
        .getRange(row, 10)
        .setBackground("#E9F7EF")
        .setFontWeight("bold");
      continue;
    }

    if (!(putoutD instanceof Date)) {
      currentSheet
        .getRange(row, 9)
        .setBackground("#F8D7DA")
        .setFontWeight("bold");
      continue;
    }

    // Case 4: putout date < send date → RED
    currentSheet
      .getRange(row, 9)
      .setBackground("#F8D7DA")
      .setFontWeight("bold");
  }
}
