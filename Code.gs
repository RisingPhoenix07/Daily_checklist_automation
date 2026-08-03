// -----------------------------
// Dashboard setup + automation
// Debugged + optimized
// -----------------------------
// You can rename segment or task names from here or from the sheet as well :)

var SEGMENT_NAMES = [
  "Algorithms Lab",
  "Circuit Design",
  "Math Practice",
  "Psychology Notes",
  "Guitar Training",
  "Fitness Routine",
  "Personal Growth"
];

// Every segment gets this same set of 7 tasks each day.
var TASK_NAMES = [
  "Solve 5 problems",
  "Summarize lecture notes",
  "Practice 2 circuits",
  "Write reflection journal",
  "Play scales 15 min",
  "Do 20 pushups",
  "Read 10 pages"
];

var NUM_SEGMENTS = SEGMENT_NAMES.length;
var TASKS_PER_SEGMENT = TASK_NAMES.length;

// Daily Checklist row layout is now fully deterministic: each segment is
// 1 header row (name + % complete in the same row) followed by its task
// rows — no separate "% Complete" row anymore, and no need to scan column A
// for the word "complete" to figure out where a segment ends.
function getSegmentLayout_() {
  var segments = [];
  var row = 2; // row 1 is the sheet header
  for (var s = 0; s < NUM_SEGMENTS; s++) {
    var headerRow = row;
    var startRow = headerRow + 1;
    var endRow = startRow + TASKS_PER_SEGMENT - 1;
    segments.push({ name: SEGMENT_NAMES[s], headerRow: headerRow, startRow: startRow, endRow: endRow });
    row = endRow + 1;
  }
  return segments;
}

// -----------------------------
// Setup
// -----------------------------
function setupDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  setupDailyChecklist_(ss);
  setupWeeklyProgress_(ss);
  setupDailyCompletion_(ss);

  setupGradientFormatting();
  refreshWeeklyChart();
  refreshDailyChart();

  ensureDailyTrigger_();
}

function setupDailyChecklist_(ss) {
  var sheet = ss.getSheetByName("Daily Checklist") || ss.insertSheet("Daily Checklist");

  // Full clear including data validation (checkboxes), so re-runs are clean.
  sheet.clear();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();

  var segments = getSegmentLayout_();
  var values = [];
  var boldRows = [1]; // sheet header row is bold too
  var checkboxRanges = []; // {startRow, numRows}

  // Sheet header row
  values.push([
    "Segment", "Daily Focus", "Comment",
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM")
  ]);

  segments.forEach(function (seg) {
    // Header row: segment name in col A, % complete formula goes in col D below
    values.push([seg.name, "", "", ""]);
    boldRows.push(seg.headerRow);

    for (var t = 0; t < TASKS_PER_SEGMENT; t++) {
      values.push(["", TASK_NAMES[t], "", ""]); // checkbox applied below
    }
    checkboxRanges.push({ startRow: seg.startRow, numRows: TASKS_PER_SEGMENT });
  });

  // 1. Write all text content in a single batched call
  sheet.getRange(1, 1, values.length, 4).setValues(values);

  // 2. Alignment: column A left-aligned (long segment names would otherwise
  //    overflow both directions when centered, clipping behind the row-number
  //    gutter on the left); columns B-D stay centered.
  sheet.getRange(1, 1, values.length, 1).setHorizontalAlignment("left");
  sheet.getRange(1, 2, values.length, 3).setHorizontalAlignment("center");

  // Widen column A so segment names have room and don't need to overflow at all
  sheet.autoResizeColumn(1);
  sheet.setColumnWidth(1, sheet.getColumnWidth(1) + 20);

  // 3. Bold header rows
  boldRows.forEach(function (r) {
    sheet.getRange(r, 1, 1, 4).setFontWeight("bold");
  });

  // 4. Insert checkboxes per segment (contiguous ranges -> 7 calls, not 49)
  checkboxRanges.forEach(function (range) {
    sheet.getRange(range.startRow, 4, range.numRows, 1).insertCheckboxes();
  });

  // 5. % complete formula lives in the segment's own header row now
  segments.forEach(function (seg) {
    var range = "D" + seg.startRow + ":D" + seg.endRow;
    var cell = sheet.getRange(seg.headerRow, 4);
    cell.setFormula("=COUNTIF(" + range + ",TRUE)/COUNTA(" + range + ")");
    cell.setNumberFormat("0%");
    cell.setHorizontalAlignment("center");
  });
}

function setupWeeklyProgress_(ss) {
  var sheet = ss.getSheetByName("Weekly Progress") || ss.insertSheet("Weekly Progress");
  sheet.clear();

  var values = [["Segment"]];
  for (var s = 0; s < NUM_SEGMENTS; s++) {
    values.push([SEGMENT_NAMES[s]]);
  }
  values.push(["Weekly Average"]);

  var range = sheet.getRange(1, 1, values.length, 1);
  range.setValues(values);
  range.setHorizontalAlignment("left");
  sheet.autoResizeColumn(1);
  sheet.setColumnWidth(1, sheet.getColumnWidth(1) + 20);
  sheet.getRange(1, 1).setFontWeight("bold");
  sheet.getRange(values.length, 1).setFontWeight("bold");
}

function setupDailyCompletion_(ss) {
  var sheet = ss.getSheetByName("Daily Completion") || ss.insertSheet("Daily Completion");
  sheet.clear();

  var header = sheet.getRange(1, 1, 1, 2);
  header.setValues([["Date", "Overall %"]]);
  header.setFontWeight("bold").setHorizontalAlignment("center");
}

// -----------------------------
// Add daily column + logging
// -----------------------------

// Shared core: creates one day's column across all three sheets.
// randomizeCheckboxes=true is used only by seedRandomWeekData() for demo/test data;
// the real daily trigger always calls this with false (fresh, unchecked boxes).
function createDayEntry_(dateLabel, randomizeCheckboxes) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dailySheet = ss.getSheetByName("Daily Checklist");
  if (!dailySheet) throw new Error("Daily Checklist sheet not found.");
  var weeklySheet = ss.getSheetByName("Weekly Progress");
  if (!weeklySheet) throw new Error("Weekly Progress sheet not found.");

  var lastCol = dailySheet.getLastColumn();
  var dateString = dateLabel;

  // Insert new date column in Daily Checklist
  dailySheet.insertColumnAfter(lastCol);
  var newCol = lastCol + 1;
  dailySheet.getRange(1, newCol).setValue(dateString).setHorizontalAlignment("center");

  // Row layout is fixed/known up front now (no scanning needed)
  var segments = getSegmentLayout_();

  // Batch checkboxes + percent formulas, once per segment
  segments.forEach(function (seg) {
    var taskCount = seg.endRow - seg.startRow + 1;
    var percentCell = dailySheet.getRange(seg.headerRow, newCol);

    var checkRange = dailySheet.getRange(seg.startRow, newCol, taskCount, 1);
    checkRange.insertCheckboxes();

    if (randomizeCheckboxes) {
      // Weighted ~70% checked so seeded weeks look like plausible real progress
      var randomValues = [];
      for (var k = 0; k < taskCount; k++) {
        randomValues.push([Math.random() < 0.7]);
      }
      checkRange.setValues(randomValues);
    }

    var taskRangeA1 = checkRange.getA1Notation();
    percentCell.setFormula("=COUNTIF(" + taskRangeA1 + ",TRUE)/COUNTA(" + taskRangeA1 + ")");
    percentCell.setNumberFormat("0%").setHorizontalAlignment("center");
  });

  // --- Weekly Progress: ONE new column for the whole day (bug fix) ---
  var weeklyCol = Math.max(weeklySheet.getLastColumn(), 1) + 1;
  weeklySheet.getRange(1, weeklyCol).setValue(dateString).setHorizontalAlignment("center");

  var weeklyFormulas = segments.map(function (seg) {
    return ["='" + dailySheet.getName() + "'!" + dailySheet.getRange(seg.headerRow, newCol).getA1Notation()];
  });
  var weeklyRange = weeklySheet.getRange(2, weeklyCol, weeklyFormulas.length, 1);
  weeklyRange.setFormulas(weeklyFormulas);
  weeklyRange.setNumberFormat("0%").setHorizontalAlignment("center");

  var weeklyAvgCell = weeklySheet.getRange(9, weeklyCol);
  weeklyAvgCell.setFormula(
    "=AVERAGE(" + weeklySheet.getRange(2, weeklyCol, segments.length).getA1Notation() + ")"
  ).setNumberFormat("0%").setHorizontalAlignment("center");

  // --- Daily Completion log ---
  var dailyCompletion = ss.getSheetByName("Daily Completion") || ss.insertSheet("Daily Completion");
  if (dailyCompletion.getLastRow() === 0) {
    dailyCompletion.getRange(1, 1, 1, 2).setValues([["Date", "Overall %"]])
      .setFontWeight("bold").setHorizontalAlignment("center");
  }
  var newRow = dailyCompletion.getLastRow() + 1;
  dailyCompletion.getRange(newRow, 1).setValue(dateString);

  var overallCell = dailyCompletion.getRange(newRow, 2);
  // Single source of truth: reference the Weekly Average cell we just wrote,
  // instead of re-deriving the average a second time.
  overallCell.setFormula("='" + weeklySheet.getName() + "'!" + weeklySheet.getRange(9, weeklyCol).getA1Notation());
  overallCell.setNumberFormat("0%").setHorizontalAlignment("center");
}

// Called by the daily 6 AM trigger. Always creates fresh, unchecked boxes.
function addDailyColumn() {
  var dateString = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM");
  createDayEntry_(dateString, false);

  setupGradientFormatting();
  refreshWeeklyChart();
  refreshDailyChart();
}

// Run this MANUALLY, once, to populate the last 7 days with random test data
// so you can see the gradient formatting and both charts working end to end.
// Safe to run on a freshly-setupDashboard() sheet. Re-running setupDashboard()
// afterward wipes the seeded data along with everything else, as usual.
function seedRandomWeekData() {
  var tz = Session.getScriptTimeZone();

  for (var daysAgo = 6; daysAgo >= 0; daysAgo--) {
    var date = new Date();
    date.setDate(date.getDate() - daysAgo);
    var label = Utilities.formatDate(date, tz, "dd-MMM");
    createDayEntry_(label, true);
  }

  setupGradientFormatting();
  refreshWeeklyChart();
  refreshDailyChart();
}

// -----------------------------
// Gradient formatting
// -----------------------------
function setupGradientFormatting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dailySheet = ss.getSheetByName("Daily Checklist");
  if (!dailySheet) return;

  var lastCol = dailySheet.getLastColumn();
  var lastRow = dailySheet.getLastRow();
  if (lastRow < 3 || lastCol < 4) return;

  var dateRange = dailySheet.getRange(2, 4, lastRow - 1, lastCol - 3);

  var rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpointWithValue("#FF0000", SpreadsheetApp.InterpolationType.NUMBER, "0")
      .setGradientMidpointWithValue("#FFFF00", SpreadsheetApp.InterpolationType.NUMBER, "0.5")
      .setGradientMaxpointWithValue("#00FFFF", SpreadsheetApp.InterpolationType.NUMBER, "1")
      .setRanges([dateRange])
      .build()
  ];
  dailySheet.setConditionalFormatRules(rules);
}

// -----------------------------
// Chart refresh functions
// -----------------------------
function refreshDailyChart() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dailyCompletion = ss.getSheetByName("Daily Completion");
  if (!dailyCompletion) return;

  var charts = dailyCompletion.getCharts();
  charts.forEach(function (chart) { dailyCompletion.removeChart(chart); });

  var lastRow = dailyCompletion.getLastRow();
  if (lastRow < 2) return; // no data yet

  var chart = dailyCompletion.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dailyCompletion.getRange("A1:B" + lastRow))
    .setPosition(2, 4, 0, 0)
    .setOption("title", "Daily Completion")
    .setOption("vAxis.format", "percent")
    .setOption("legend", { position: "none" })
    .build();
  dailyCompletion.insertChart(chart);
}

function refreshWeeklyChart() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var weeklySheet = ss.getSheetByName("Weekly Progress");
  if (!weeklySheet) return;

  var lastRow = weeklySheet.getLastRow();
  var lastCol = weeklySheet.getLastColumn();
  if (lastRow < 9 || lastCol < 2) return; // need segments + at least one date

  var charts = weeklySheet.getCharts();
  charts.forEach(function (chart) { weeklySheet.removeChart(chart); });

  // Bug fix: previously row `lastRow` (Weekly Average) was included in the
  // first range AND added again explicitly, duplicating that series. Now the
  // two ranges are non-overlapping: segments block, then the average row.
  var chart = weeklySheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(weeklySheet.getRange(1, 1, lastRow - 1, lastCol)) // header + segments only
    .addRange(weeklySheet.getRange(lastRow, 1, 1, lastCol))     // weekly average overlay
    .setPosition(2, 4, 0, 0)
    .setOption("title", "Weekly Progress with Average")
    .setOption("vAxis.format", "percent")
    .setOption("curveType", "function")
    .setOption("series", {
      7: { color: "black", lineDashStyle: [4, 4], labelInLegend: "Weekly Average" }
    })
    .build();
  weeklySheet.insertChart(chart);
}

// -----------------------------
// Trigger helpers
// -----------------------------
function deleteTriggersByHandler_(handlerName) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function ensureDailyTrigger_() {
  var exists = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === "addDailyColumn";
  });
  if (!exists) {
    ScriptApp.newTrigger("addDailyColumn")
      .timeBased()
      .everyDays(1)
      .atHour(6) // runs daily at 6 AM (script timezone)
      .create();
  }
}

function createTestTrigger() {
  // Delete existing addDailyColumn triggers first, then create a fast one for testing.
  deleteTriggersByHandler_("addDailyColumn");
  ScriptApp.newTrigger("addDailyColumn")
    .timeBased()
    .everyMinutes(1)
    .create();
}

function deleteAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    ScriptApp.deleteTrigger(t);
  });
}
