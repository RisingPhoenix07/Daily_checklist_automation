// =====================================================================
// Dashboard setup + automation  (reviewed + extended)
//
// NEW: Dashboard menu -> "Edit segments & tasks..." lets you rename,
// add, remove or reorder segments and tasks at any time. Existing
// checkbox history and comments are kept (matched by position).
// =====================================================================

var SHEET_DAILY = "Daily Checklist";
var SHEET_WEEKLY = "Weekly Progress";
var SHEET_COMPLETION = "Daily Completion";
var SHEET_CHARTDATA = "Chart Data"; // helper sheet feeding the weekly chart (auto-managed)
var FIRST_DATE_COL = 4; // Daily Checklist: A=Segment, B=Focus, C=Comment, D+ = dates
var CONFIG_KEY = "DASHBOARD_CONFIG";
var SHEET_HOME = "🔥 Rising Phoenix"; // welcome / stats page, always the first tab
var HOME_TARGET_A1 = "E21";           // editable "streak target" cell on the home page
var DEFAULT_STREAK_TARGET = 0.7;      // a day counts toward the streak at 70%+ completion

// Original motivational lines. A new one is picked every time the file opens,
// each morning, and from Dashboard > New motivational quote. Edit freely.
var QUOTES = [
  "Ashes are just the beginning. Rise anyway.",
  "Small steps, repeated daily, build a fire that cannot be put out.",
  "You don't need a perfect day. You need a day you showed up.",
  "Burn the excuses, keep the embers of discipline.",
  "Every checked box is a vote for the person you're becoming.",
  "Progress is quiet. Consistency is loud.",
  "Yesterday's ashes are today's fuel.",
  "Rise a little higher than you did yesterday.",
  "Discipline is choosing what you want most over what you want now.",
  "A streak isn't about perfection; it's about never letting a bad day become a bad week.",
  "Start where you are. Use what you have. Burn bright anyway.",
  "Motivation lights the spark; habit keeps the flame.",
  "One honest effort today beats a hundred perfect plans.",
  "You have survived every hard day so far. Today is no different.",
  "Be the phoenix: fall, burn, rebuild, rise.",
  "Your future self is watching what you do right now.",
  "Tiny wins stack into a transformation.",
  "Don't wait to feel ready. Move, and readiness will catch up.",
  "A missed day is a pause, not a verdict.",
  "Show up tired, show up late, but show up.",
  "Heat reveals what you're made of.",
  "Build the habit and the goal becomes a by-product.",
  "The hardest rep is the one you skip.",
  "Do it for the version of you waiting on the other side.",
  "Comparison is noise; your own progress is the signal.",
  "Every sunrise is a fresh set of boxes to check.",
  "Courage is rebuilding after the fire, not avoiding it.",
  "Make today's effort something tomorrow-you will thank you for.",
  "Repetition quietly writes your story.",
  "Rise. Check. Repeat."
];

// Defaults, used only the FIRST time (before you save a custom config).
// After that, change names/counts from the Dashboard menu instead.
var DEFAULT_SEGMENTS = ["Segment 1", "Segment 2", "Segment 3", "Segment 4", "Segment 5", "Segment 6", "Segment 7"];
var DEFAULT_TASKS = ["Task 1", "Task 2", "Task 3", "Task 4", "Task 5", "Task 6", "Task 7"];

// -----------------------------
// Menu
// -----------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Dashboard")
    .addItem("Edit segments & tasks...", "showConfigDialog")
    .addItem("Add today's column now", "addDailyColumn")
    .addItem("Refresh charts", "refreshCharts")
    .addItem("New motivational quote", "newQuote")
    .addItem("Rebuild home page", "rebuildHome")
    .addSeparator()
    .addItem("Run setup (erases history)", "setupDashboard")
    .addItem("Seed 7 days of test data", "seedRandomWeekData")
    .addItem("Reset names to defaults", "resetToDefaults")
    .addToUi();

  // Fresh quote each time the spreadsheet is opened (never block the menu on it).
  try { refreshQuote_(); } catch (e) {}
}

// -----------------------------
// Config (stored in document properties, so it survives edits)
// -----------------------------
function getConfig_() {
  var raw = PropertiesService.getDocumentProperties().getProperty(CONFIG_KEY);
  if (raw) {
    try {
      var c = JSON.parse(raw);
      if (c.segments && c.segments.length && c.tasks && c.tasks.length) return c;
    } catch (e) {}
  }
  return { segments: DEFAULT_SEGMENTS.slice(), tasks: DEFAULT_TASKS.slice() };
}

function saveConfig_(cfg) {
  PropertiesService.getDocumentProperties().setProperty(CONFIG_KEY, JSON.stringify(cfg));
}

// Row layout of Daily Checklist: 1 header row per segment (name + % complete)
// followed by that segment's task rows.
function getSegmentLayout_(cfg) {
  cfg = cfg || getConfig_();
  var segments = [];
  var row = 2; // row 1 is the sheet header
  for (var s = 0; s < cfg.segments.length; s++) {
    var headerRow = row;
    var startRow = headerRow + 1;
    var endRow = startRow + cfg.tasks.length - 1;
    segments.push({ name: cfg.segments[s], headerRow: headerRow, startRow: startRow, endRow: endRow });
    row = endRow + 1;
  }
  return segments;
}

// Weekly Progress: row 1 header, rows 2..N+1 segments, row N+2 average.
function weeklyAvgRow_(cfg) {
  return cfg.segments.length + 2;
}

// -----------------------------
// Small helpers
// -----------------------------
function colLetter_(n) {
  var s = "";
  while (n > 0) {
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function resetSheet_(sheet) {
  sheet.getCharts().forEach(function (c) { sheet.removeChart(c); });
  sheet.clear();
  sheet.clearConditionalFormatRules();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
}

function ensureSize_(sheet, rows, cols) {
  if (sheet.getMaxRows() < rows) sheet.insertRowsAfter(sheet.getMaxRows(), rows - sheet.getMaxRows());
  if (sheet.getMaxColumns() < cols) sheet.insertColumnsAfter(sheet.getMaxColumns(), cols - sheet.getMaxColumns());
}

function todayLabel_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM");
}

// -----------------------------
// Core builder: (re)creates all three sheets from config + history.
//   dates:    array of date labels (strings)
//   history:  history[dateIdx][segIdx][taskIdx] = true/false   (optional)
//   comments: comments[segIdx][taskIdx] = string               (optional)
// Everything is written in a few batched calls.
// -----------------------------
function rebuildDashboard_(cfg, dates, history, comments) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var layout = getSegmentLayout_(cfg);
  var nSeg = cfg.segments.length;
  var nDates = dates.length;
  var avgRow = weeklyAvgRow_(cfg);

  // The home sheet must exist before any formula refers to it (the streak
  // column below reads the editable target from it). Keep the user's target
  // across rebuilds; the page itself is filled in at the very end.
  var streakTarget = getHomeTarget_(ss);
  getHomeSheet_(ss);

  // ---------- Daily Checklist ----------
  var daily = getOrCreateSheet_(ss, SHEET_DAILY);
  resetSheet_(daily);

  var totalRows = layout[layout.length - 1].endRow;
  var width = Math.max(FIRST_DATE_COL - 1 + nDates, FIRST_DATE_COL - 1);
  ensureSize_(daily, totalRows, width);

  var grid = [];
  for (var r = 0; r < totalRows; r++) {
    var blank = [];
    for (var c = 0; c < width; c++) blank.push("");
    grid.push(blank);
  }
  grid[0][0] = "Segment";
  grid[0][1] = "Daily Focus";
  grid[0][2] = "Comment";
  for (var d = 0; d < nDates; d++) grid[0][FIRST_DATE_COL - 1 + d] = dates[d];

  layout.forEach(function (seg, s) {
    grid[seg.headerRow - 1][0] = seg.name;
    for (var d2 = 0; d2 < nDates; d2++) {
      var L = colLetter_(FIRST_DATE_COL + d2);
      var rng = L + seg.startRow + ":" + L + seg.endRow;
      grid[seg.headerRow - 1][FIRST_DATE_COL - 1 + d2] = "=COUNTIF(" + rng + ",TRUE)/COUNTA(" + rng + ")";
    }
    for (var t = 0; t < cfg.tasks.length; t++) {
      var rowIdx = seg.startRow - 1 + t;
      grid[rowIdx][1] = cfg.tasks[t];
      grid[rowIdx][2] = (comments && comments[s] && comments[s][t]) || "";
      for (var d3 = 0; d3 < nDates; d3++) {
        grid[rowIdx][FIRST_DATE_COL - 1 + d3] =
          !!(history && history[d3] && history[d3][s] && history[d3][s][t]);
      }
    }
  });

  // Date headers as plain text so Sheets doesn't turn "15-Oct" into a real date.
  if (nDates > 0) daily.getRange(1, FIRST_DATE_COL, 1, nDates).setNumberFormat("@");
  daily.getRange(1, 1, totalRows, width).setValues(grid);

  daily.getRange(1, 1, totalRows, 1).setHorizontalAlignment("left");
  daily.getRange(1, 2, totalRows, width - 1).setHorizontalAlignment("center");
  daily.getRange(1, 1, 1, width).setFontWeight("bold");

  layout.forEach(function (seg) {
    daily.getRange(seg.headerRow, 1, 1, width).setFontWeight("bold");
    if (nDates > 0) {
      daily.getRange(seg.headerRow, FIRST_DATE_COL, 1, nDates).setNumberFormat("0%");
      daily.getRange(seg.startRow, FIRST_DATE_COL, cfg.tasks.length, nDates).insertCheckboxes();
    }
  });

  daily.autoResizeColumn(1);
  daily.setColumnWidth(1, daily.getColumnWidth(1) + 20);
  daily.autoResizeColumn(2);
  daily.setColumnWidth(2, daily.getColumnWidth(2) + 20);

  // ---------- Weekly Progress ----------
  var weekly = getOrCreateSheet_(ss, SHEET_WEEKLY);
  resetSheet_(weekly);
  var wWidth = 1 + nDates;
  ensureSize_(weekly, avgRow, wWidth);

  var wGrid = [];
  var wHead = ["Segment"];
  for (var d4 = 0; d4 < nDates; d4++) wHead.push(dates[d4]);
  wGrid.push(wHead);

  layout.forEach(function (seg) {
    // Labels link back to Daily Checklist, so renaming there flows through.
    var row = ["='" + SHEET_DAILY + "'!A" + seg.headerRow];
    for (var d5 = 0; d5 < nDates; d5++) {
      row.push("='" + SHEET_DAILY + "'!" + colLetter_(FIRST_DATE_COL + d5) + seg.headerRow);
    }
    wGrid.push(row);
  });

  var avgLine = ["Weekly Average"];
  for (var d6 = 0; d6 < nDates; d6++) {
    var WL = colLetter_(2 + d6);
    avgLine.push("=AVERAGE(" + WL + "2:" + WL + (nSeg + 1) + ")");
  }
  wGrid.push(avgLine);

  if (nDates > 0) weekly.getRange(1, 2, 1, nDates).setNumberFormat("@");
  weekly.getRange(1, 1, avgRow, wWidth).setValues(wGrid);
  if (nDates > 0) {
    weekly.getRange(2, 2, nSeg + 1, nDates).setNumberFormat("0%");
    weekly.getRange(1, 2, avgRow, nDates).setHorizontalAlignment("center");
  }
  weekly.getRange(1, 1, avgRow, 1).setHorizontalAlignment("left");
  weekly.getRange(1, 1, 1, wWidth).setFontWeight("bold");
  weekly.getRange(avgRow, 1, 1, wWidth).setFontWeight("bold");
  weekly.autoResizeColumn(1);
  weekly.setColumnWidth(1, weekly.getColumnWidth(1) + 20);

  // ---------- Daily Completion ----------
  var comp = getOrCreateSheet_(ss, SHEET_COMPLETION);
  resetSheet_(comp);
  ensureSize_(comp, nDates + 1, 3);
  var cGrid = [["Date", "Overall %", "Streak count"]];
  for (var d7 = 0; d7 < nDates; d7++) {
    cGrid.push([
      dates[d7],
      "='" + SHEET_WEEKLY + "'!" + colLetter_(2 + d7) + avgRow,
      streakFormula_(2 + d7)
    ]);
  }
  comp.getRange(1, 1, nDates + 1, 1).setNumberFormat("@");
  comp.getRange(1, 1, nDates + 1, 3).setValues(cGrid);
  if (nDates > 0) comp.getRange(2, 2, nDates, 1).setNumberFormat("0%");
  comp.getRange(1, 1, nDates + 1, 3).setHorizontalAlignment("center");
  comp.getRange(1, 1, 1, 3).setFontWeight("bold");

  // ---------- Home ----------
  buildHome_(ss, cfg, streakTarget);
}

function refreshAll_() {
  setupGradientFormatting();
  refreshDailyChart();  // simplest chart first, so it still appears if the weekly one hits a problem
  refreshWeeklyChart();
}

// Public wrapper for the menu (menu items can't call names ending in "_").
function refreshCharts() {
  refreshAll_();
}

// -----------------------------
// Setup (fresh start)
// -----------------------------
function setupDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var existing = ss.getSheetByName(SHEET_DAILY);
  if (existing && existing.getLastColumn() > FIRST_DATE_COL) {
    var ui = SpreadsheetApp.getUi();
    var answer = ui.alert(
      "Reset dashboard?",
      "This erases ALL existing checklist history and starts again with one column for today. " +
        "(To change segments/tasks without losing history, use Dashboard > Edit segments & tasks.)",
      ui.ButtonSet.YES_NO
    );
    if (answer !== ui.Button.YES) return;
  }

  var cfg = getConfig_();
  saveConfig_(cfg);
  // Today's column is created as a real day entry (the old version made a
  // stray column that Weekly Progress / Daily Completion never tracked).
  rebuildDashboard_(cfg, [todayLabel_()], [], null);
  ensureDailyTrigger_(); // before the charts, so a chart problem can never skip it
  refreshAll_();
}

function resetToDefaults() {
  PropertiesService.getDocumentProperties().deleteProperty(CONFIG_KEY);
  setupDashboard();
}

// -----------------------------
// Edit segments & tasks (rename / add / remove / reorder)
// -----------------------------
function showConfigDialog() {
  var cfg = getConfig_();
  var esc = function (s) { return JSON.stringify(s).replace(/</g, "\\u003c"); };

  var html =
    '<style>' +
    'body{font-family:Arial,sans-serif;margin:12px;font-size:13px}' +
    'textarea{width:100%;height:150px;box-sizing:border-box;font-size:13px}' +
    'button{margin-top:10px;padding:6px 16px}#msg{color:#b00;margin-left:8px}' +
    '.hint{color:#666;font-size:12px;margin-top:6px}' +
    '</style>' +
    '<b>Segments</b> (one per line)<br><textarea id="seg"></textarea><br><br>' +
    '<b>Tasks</b> (one per line, same set for every segment)<br><textarea id="task"></textarea>' +
    '<div class="hint">History and comments are kept and matched by position: ' +
    'the 3rd task stays the 3rd task, even if renamed. Removing lines at the end ' +
    'drops that data; new lines start unchecked on past days.</div>' +
    '<button id="save">Save</button><span id="msg"></span>' +
    '<script>' +
    'document.getElementById("seg").value=' + esc(cfg.segments.join("\n")) + ';' +
    'document.getElementById("task").value=' + esc(cfg.tasks.join("\n")) + ';' +
    'document.getElementById("save").onclick=function(){' +
    'var b=this;b.disabled=true;' +
    'document.getElementById("msg").textContent="Saving...";' +
    'google.script.run' +
    '.withSuccessHandler(function(){google.script.host.close();})' +
    '.withFailureHandler(function(e){b.disabled=false;document.getElementById("msg").textContent=e.message;})' +
    '.applyDashboardConfig(document.getElementById("seg").value,document.getElementById("task").value);' +
    '};' +
    '</script>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(420).setHeight(470),
    "Edit segments & tasks"
  );
}

function parseLines_(text) {
  return String(text || "")
    .split(/\r?\n/)
    .map(function (s) { return s.trim(); })
    .filter(function (s) { return s.length > 0; });
}

// Public (no trailing underscore) so the dialog can call it. You can also
// call updateDashboardConfig(["Work","Health"], ["Plan","Do","Review"]) directly.
function applyDashboardConfig(segmentsText, tasksText) {
  return updateDashboardConfig(parseLines_(segmentsText), parseLines_(tasksText));
}

function updateDashboardConfig(newSegments, newTasks) {
  if (!newSegments.length || !newTasks.length) {
    throw new Error("You need at least one segment and one task.");
  }
  if (newSegments.length > 50 || newTasks.length > 50) {
    throw new Error("Please keep it to 50 segments and 50 tasks or fewer.");
  }

  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var oldCfg = getConfig_();
    var newCfg = { segments: newSegments, tasks: newTasks };
    var daily = ss.getSheetByName(SHEET_DAILY);

    // No dashboard yet: just save and build it.
    if (!daily) {
      saveConfig_(newCfg);
      rebuildDashboard_(newCfg, [todayLabel_()], [], null);
      refreshAll_();
      ensureDailyTrigger_();
      return "Created";
    }

    // Read existing history + comments using the OLD layout.
    var oldLayout = getSegmentLayout_(oldCfg);
    var oldLastRow = oldLayout[oldLayout.length - 1].endRow;
    if (daily.getLastRow() !== oldLastRow) {
      throw new Error(
        "The Daily Checklist layout doesn't match the saved settings (rows were added or removed by hand). " +
          "Undo those manual row changes, or run setup to start fresh."
      );
    }

    var lastCol = daily.getLastColumn();
    var nDates = Math.max(lastCol - FIRST_DATE_COL + 1, 0);
    var dates = [];
    var history = [];
    if (nDates > 0) {
      dates = daily.getRange(1, FIRST_DATE_COL, 1, nDates).getDisplayValues()[0];
      var data = daily.getRange(1, FIRST_DATE_COL, oldLastRow, nDates).getValues();
      for (var d = 0; d < nDates; d++) {
        history[d] = [];
        oldLayout.forEach(function (seg, s) {
          history[d][s] = [];
          for (var t = 0; t < oldCfg.tasks.length; t++) {
            history[d][s][t] = data[seg.startRow - 1 + t][d] === true;
          }
        });
      }
    }
    var commentCol = daily.getRange(1, 3, oldLastRow, 1).getValues();
    var comments = oldLayout.map(function (seg) {
      var arr = [];
      for (var t = 0; t < oldCfg.tasks.length; t++) arr.push(commentCol[seg.startRow - 1 + t][0]);
      return arr;
    });

    saveConfig_(newCfg);
    rebuildDashboard_(newCfg, dates, history, comments);
    refreshAll_();
    ss.toast(newSegments.length + " segments x " + newTasks.length + " tasks saved.", "Dashboard", 5);
    return "OK";
  } finally {
    lock.releaseLock();
  }
}

// -----------------------------
// Daily column (called by the 6 AM trigger)
// -----------------------------
// Appends ONE day across all three sheets. Returns false if that date
// already exists (protects against double trigger runs).
function addDayEntry_(dateLabel) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = getConfig_();
  var layout = getSegmentLayout_(cfg);
  var avgRow = weeklyAvgRow_(cfg);

  var daily = ss.getSheetByName(SHEET_DAILY);
  var weekly = ss.getSheetByName(SHEET_WEEKLY);
  var comp = ss.getSheetByName(SHEET_COMPLETION);
  if (!daily) throw new Error("Daily Checklist sheet not found. Run setup first.");
  if (!weekly) throw new Error("Weekly Progress sheet not found. Run setup first.");
  if (!comp) throw new Error("Daily Completion sheet not found. Run setup first.");

  var lastCol = daily.getLastColumn();
  var existing = lastCol >= FIRST_DATE_COL
    ? daily.getRange(1, FIRST_DATE_COL, 1, lastCol - FIRST_DATE_COL + 1).getDisplayValues()[0]
    : [];
  if (existing.indexOf(dateLabel) !== -1) return false;

  // Date index i -> Daily col 4+i, Weekly col 2+i, Completion row 2+i
  var idx = existing.length;
  var dCol = FIRST_DATE_COL + idx;
  var wCol = 2 + idx;
  var cRow = 2 + idx;

  // Daily Checklist
  ensureSize_(daily, daily.getMaxRows(), dCol);
  daily.getRange(1, dCol).setNumberFormat("@").setValue(dateLabel)
    .setHorizontalAlignment("center").setFontWeight("bold");
  var L = colLetter_(dCol);
  layout.forEach(function (seg) {
    var n = seg.endRow - seg.startRow + 1;
    daily.getRange(seg.startRow, dCol, n, 1).insertCheckboxes();
    var rng = L + seg.startRow + ":" + L + seg.endRow;
    daily.getRange(seg.headerRow, dCol)
      .setFormula("=COUNTIF(" + rng + ",TRUE)/COUNTA(" + rng + ")")
      .setNumberFormat("0%").setHorizontalAlignment("center").setFontWeight("bold");
  });

  // Weekly Progress (one column for the whole day)
  ensureSize_(weekly, avgRow, wCol);
  weekly.getRange(1, wCol).setNumberFormat("@").setValue(dateLabel)
    .setHorizontalAlignment("center").setFontWeight("bold");
  var formulas = layout.map(function (seg) {
    return ["='" + SHEET_DAILY + "'!" + L + seg.headerRow];
  });
  weekly.getRange(2, wCol, formulas.length, 1).setFormulas(formulas)
    .setNumberFormat("0%").setHorizontalAlignment("center");
  var WL = colLetter_(wCol);
  weekly.getRange(avgRow, wCol)
    .setFormula("=AVERAGE(" + WL + "2:" + WL + (cfg.segments.length + 1) + ")")
    .setNumberFormat("0%").setHorizontalAlignment("center").setFontWeight("bold");

  // Daily Completion
  ensureSize_(comp, cRow, 3);
  comp.getRange(cRow, 1).setNumberFormat("@").setValue(dateLabel).setHorizontalAlignment("center");
  comp.getRange(cRow, 2).setFormula("='" + SHEET_WEEKLY + "'!" + WL + avgRow)
    .setNumberFormat("0%").setHorizontalAlignment("center");
  comp.getRange(cRow, 3).setFormula(streakFormula_(cRow)).setHorizontalAlignment("center");

  return true;
}

function addDailyColumn() {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    if (addDayEntry_(todayLabel_())) {
      refreshAll_();
      refreshQuote_(); // new day, new quote
    }
  } finally {
    lock.releaseLock();
  }
}

// Manual: wipes everything and fills the last 7 days with random data
// (~70% checked) so you can preview the formatting and charts.
function seedRandomWeekData() {
  var cfg = getConfig_();
  var tz = Session.getScriptTimeZone();
  var dates = [];
  var history = [];

  for (var daysAgo = 6; daysAgo >= 0; daysAgo--) {
    var date = new Date();
    date.setDate(date.getDate() - daysAgo);
    dates.push(Utilities.formatDate(date, tz, "dd-MMM"));
    var day = [];
    for (var s = 0; s < cfg.segments.length; s++) {
      var row = [];
      for (var t = 0; t < cfg.tasks.length; t++) row.push(Math.random() < 0.7);
      day.push(row);
    }
    history.push(day);
  }

  rebuildDashboard_(cfg, dates, history, null);
  refreshAll_();
}

// -----------------------------
// Gradient formatting
// -----------------------------
function setupGradientFormatting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var daily = ss.getSheetByName(SHEET_DAILY);
  if (!daily) return;

  var lastCol = daily.getLastColumn();
  var lastRow = daily.getLastRow();
  if (lastRow < 3 || lastCol < FIRST_DATE_COL) return;

  var dateRange = daily.getRange(2, FIRST_DATE_COL, lastRow - 1, lastCol - FIRST_DATE_COL + 1);
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue("#FF0000", SpreadsheetApp.InterpolationType.NUMBER, "0")
    .setGradientMidpointWithValue("#FFFF00", SpreadsheetApp.InterpolationType.NUMBER, "0.5")
    .setGradientMaxpointWithValue("#00FFFF", SpreadsheetApp.InterpolationType.NUMBER, "1")
    .setRanges([dateRange])
    .build();
  daily.setConditionalFormatRules([rule]);
}

// -----------------------------
// Charts
// -----------------------------
function refreshDailyChart() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var comp = ss.getSheetByName(SHEET_COMPLETION);
  if (!comp) return;

  comp.getCharts().forEach(function (c) { comp.removeChart(c); });

  var lastRow = comp.getLastRow();
  if (lastRow < 2) return;

  var chart = comp.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(comp.getRange("A1:B" + lastRow))
    .setPosition(2, 4, 0, 0)
    .setOption("title", "Daily Completion")
    .setOption("vAxis.format", "percent")
    .setOption("legend", { position: "none" })
    .build();
  comp.insertChart(chart);
}

function refreshWeeklyChart() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var weekly = ss.getSheetByName(SHEET_WEEKLY);
  if (!weekly) return;

  var cfg = getConfig_();
  var avgRow = weeklyAvgRow_(cfg);
  var lastCol = weekly.getLastColumn();
  if (lastCol < 2) return;

  weekly.getCharts().forEach(function (c) { weekly.removeChart(c); });

  // Segments sit in rows and dates in columns, but a line chart needs dates
  // down the side and one column per segment. Apps Script has no "transpose"
  // option on the chart builder, so a helper sheet flips the table with
  // TRANSPOSE() and the chart reads from that instead.
  var previouslyActive = ss.getActiveSheet();
  var helper = ss.getSheetByName(SHEET_CHARTDATA) || ss.insertSheet(SHEET_CHARTDATA);
  helper.clear();
  var src = "'" + SHEET_WEEKLY + "'!A1:" + colLetter_(lastCol) + avgRow;
  helper.getRange(1, 1).setFormula("=TRANSPOSE(" + src + ")");
  SpreadsheetApp.flush();
  helper.getRange(1, 1, lastCol, 1).setNumberFormat("@"); // keep date labels as text
  helper.getRange(2, 2, lastCol - 1, avgRow - 1).setNumberFormat("0%");
  ss.setActiveSheet(previouslyActive); // inserting the helper sheet shouldn't steal focus

  // Series 0..N-1 are the segments; series N (last column) is the Weekly
  // Average, styled as a dashed black overlay.
  var series = {};
  series[cfg.segments.length] = { color: "black", lineDashStyle: [4, 4] };

  var chart = weekly.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(helper.getRange(1, 1, lastCol, avgRow))
    .setNumHeaders(1)
    .setPosition(avgRow + 2, 1, 0, 0)
    .setOption("title", "Weekly Progress with Average")
    .setOption("vAxis.format", "percent")
    .setOption("curveType", "function")
    .setOption("pointSize", 5) // with only one day of data a line has no length, so show dots
    .setOption("series", series)
    .build();
  weekly.insertChart(chart);
}

// -----------------------------
// Trigger helpers
// -----------------------------
function deleteTriggersByHandler_(handlerName) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === handlerName) ScriptApp.deleteTrigger(t);
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
      .atHour(6) // runs daily around 6 AM (script time zone)
      .create();
  }
}

function createTestTrigger() {
  deleteTriggersByHandler_("addDailyColumn");
  ScriptApp.newTrigger("addDailyColumn").timeBased().everyMinutes(1).create();
}

function deleteAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
}

// =====================================================================
// Home page: 🔥 Rising Phoenix
//   Motivation | Streak | Weekly Average | Quick Stats | Segment scores
// Everything on it is a live formula, except the quote (picked by script)
// and the streak target (an input cell you can edit).
// =====================================================================

function homeTargetRef_() {
  return "'" + SHEET_HOME + "'!$" + HOME_TARGET_A1.charAt(0) + "$" + HOME_TARGET_A1.substring(1);
}

// Daily Completion column C: running count of consecutive days at/above target.
function streakFormula_(row) {
  var t = homeTargetRef_();
  if (row === 2) return "=IF(B2>=" + t + ",1,0)";
  return "=IF(B" + row + ">=" + t + ",C" + (row - 1) + "+1,0)";
}

// Reuses the empty default "Sheet1" as the home page if there is one.
function getHomeSheet_(ss) {
  var home = ss.getSheetByName(SHEET_HOME);
  if (home) return home;

  var blank = ss.getSheetByName("Sheet1");
  if (blank && blank.getLastRow() === 0 && blank.getLastColumn() === 0 && ss.getSheets().length > 1) {
    blank.setName(SHEET_HOME);
    return blank;
  }
  if (blank && blank.getLastRow() === 0 && blank.getLastColumn() === 0) {
    // Only sheet in the file and it's empty: just rename it.
    blank.setName(SHEET_HOME);
    return blank;
  }
  return ss.insertSheet(SHEET_HOME);
}

function getHomeTarget_(ss) {
  var home = ss.getSheetByName(SHEET_HOME);
  if (home) {
    var v = home.getRange(HOME_TARGET_A1).getValue();
    if (typeof v === "number" && v > 0 && v <= 1) return v;
  }
  return DEFAULT_STREAK_TARGET;
}

function refreshQuote_() {
  var home = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_HOME);
  if (!home) return;
  var cell = home.getRange("B6");
  var current = String(cell.getValue());
  var q = "";
  for (var tries = 0; tries < 5; tries++) {
    q = "“" + QUOTES[Math.floor(Math.random() * QUOTES.length)] + "”";
    if (q !== current) break; // avoid showing the same line twice in a row
  }
  cell.setValue(q);
}

function newQuote() {
  refreshQuote_();
}

// Menu: redraw the home page only (keeps your streak target).
function rebuildHome() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  buildHome_(ss, getConfig_(), getHomeTarget_(ss));
}

function buildHome_(ss, cfg, target) {

  var home = getHomeSheet_(ss);
  var nSeg = cfg.segments.length;
  var firstScore = 24;
  var lastRow = firstScore + nSeg - 1;

  home.getRange(
    1,
    1,
    home.getMaxRows(),
    home.getMaxColumns()
  ).breakApart();

  resetSheet_(home);

  // Make sure sheet is large enough
  ensureSize_(home, 1000, 26);

  // Blue Theme because why not ;)
  var ACC   = "#38BDF8"; // Sky Blue
  var CARD  = "#0F172A"; // Slate Blue
  var MUTED = "#94A3B8"; // Muted Blue Gray
  var BASE  = "#020817"; // Deep Navy
  var TXT   = "#E2E8F0"; // Light Text

  var DC = "'" + SHEET_COMPLETION + "'!";

  // ==================================================
  // FULL SCREEN BACKGROUND
  // ==================================================
  home.getRange(
    1,
    1,
    home.getMaxRows(),
    home.getMaxColumns()
  )
  .setBackground(BASE)
  .setFontColor(TXT)
  .setFontFamily("Arial")
  .setVerticalAlignment("middle")
  .setHorizontalAlignment("center");

  // Wider dashboard
  home.setColumnWidth(1, 24);

  for (var c = 2; c <= 7; c++) {
    home.setColumnWidth(c, 170);
  }

  home.setColumnWidth(8, 24);

  home.setHiddenGridlines(true);
  home.setTabColor(ACC);

  // ==================================================
  // Helper block builder
  // ==================================================
  function block(a1, v, o) {

    o = o || {};

    var r = home.getRange(a1);

    if (v !== null && v !== undefined && v !== "") {

      var first = r.getCell(1, 1);

      if (typeof v === "string" && v.charAt(0) === "=") {
        first.setFormula(v);
      } else {
        first.setValue(v);
      }
    }

    if (a1.indexOf(":") > -1) {
      r.merge();
    }

    r
      .setBackground(o.bg || CARD)
      .setFontColor(o.color || TXT)
      .setFontSize(o.size || 11)
      .setFontWeight(o.bold ? "bold" : "normal")
      .setFontStyle(o.italic ? "italic" : "normal")
      .setHorizontalAlignment(o.align || "center")
      .setVerticalAlignment("middle")
      .setWrap(true);

    if (o.fmt) {
      r.setNumberFormat(o.fmt);
    }

    return r;
  }
  function header(a1, text) {
    block(a1, text, { bg: ACC, color: BASE, bold: true, size: 11, align: "left" });
  }
  function frame(a1) {
    home.getRange(a1).setBorder(true, true, true, true, false, false, BASE,
      SpreadsheetApp.BorderStyle.SOLID_THICK);
  }

  // ---- Title + welcome ----
  

block(
  "B2:G2",
  "🔥 Rising Phoenix",
  {
    bg: BASE,
    color: ACC,
    bold: true,
    size: 36
  }
);

block(
  "B3:G3",
  '="Welcome back • "&TEXT(TODAY(),"dddd, d mmmm yyyy")',
  {
    bg: BASE,
    color: MUTED,
    italic: true,
    size: 14
  }
);

home.setRowHeight(1, 20);
home.setRowHeight(2, 70);
home.setRowHeight(3, 35);
home.setRowHeight(4, 20);

  // ---- Motivation ----
  header("B5:G5", "✨  MOTIVATION");
  block("B6:G7", "", { size: 15, italic: true });
  home.setRowHeights(6, 2, 34);
  frame("B5:G7");
  home.setRowHeight(8, 14);

  // ---- Streak | Weekly Average | Today ----
  var n = "COUNT(" + DC + "B2:B)";
  header("B9:C9", "🔥  STREAK");
  header("D9:E9", "📅  WEEKLY AVERAGE");
  header("F9:G9", "🎯  TODAY");

  var streakNow =
    "IFERROR(LET(n," + n + ",last,INDEX(" + DC + "C:C,n+1),prev,IF(n>1,INDEX(" + DC + "C:C,n),0)," +
    "IF(last>0,last,prev)),0)";
  block("B10:C11", "=" + streakNow + '&IF(' + streakNow + '=1," day"," days")',
    { color: ACC, bold: true, size: 26 });
  block("D10:E11",
    "=IFERROR(AVERAGE(INDEX(" + DC + "B:B,MAX(2," + n + "-5)):INDEX(" + DC + "B:B," + n + "+1)),0)",
    { color: ACC, bold: true, size: 26, fmt: "0%" });
  block("F10:G11", "=IFERROR(INDEX(" + DC + "B:B," + n + "+1),0)",
    { color: ACC, bold: true, size: 26, fmt: "0%" });
  home.setRowHeights(10, 2, 30);

  block("B12:C12",
    '="Best: "&IFERROR(MAX(' + DC + 'C2:C),0)&" • goal "&TEXT($E$21,"0%")&"+"',
    { color: MUTED, size: 10 });
  block("D12:E12", "Last 7 days", { color: MUTED, size: 10 });
  block("F12:G12", '=IFERROR(INDEX(' + DC + 'A:A,' + n + '+1),"")', { color: MUTED, size: 10 });
  frame("B9:C12");
  frame("D9:E12");
  frame("F9:G12");
  home.setRowHeight(13, 14);

  // ---- Quick stats ----
  header("B14:G14", "📊  QUICK STATS");
  var scoreNames = "$B$" + firstScore + ":$B$" + lastRow;
  var scoreVals = "$G$" + firstScore + ":$G$" + lastRow;
  var stats = [
    ["Days tracked", "=" + n, "0"],
    ["Overall average", "=IFERROR(AVERAGE(" + DC + "B2:B),0)", "0%"],
    ["Best day",
      '=IFERROR(TEXT(MAX(' + DC + 'B2:B),"0%")&"  ("&INDEX(' + DC + 'A2:A,MATCH(MAX(' + DC + 'B2:B),' + DC + 'B2:B,0))&")","-")',
      "@"],
    ["Perfect days (100%)", "=COUNTIF(" + DC + 'B2:B,">=0.9999")', "0"],
    ["Strongest segment",
      '=IFERROR(INDEX(' + scoreNames + ',MATCH(MAX(' + scoreVals + '),' + scoreVals + ',0))&"  ("&TEXT(MAX(' + scoreVals + '),"0%")&")","-")',
      "@"],
    ["Needs attention",
      '=IFERROR(INDEX(' + scoreNames + ',MATCH(MIN(' + scoreVals + '),' + scoreVals + ',0))&"  ("&TEXT(MIN(' + scoreVals + '),"0%")&")","-")',
      "@"]
  ];
  stats.forEach(function (s, i) {
    var r = 15 + i;
    block("B" + r + ":D" + r, s[0], { align: "left", color: MUTED });
    block("E" + r + ":G" + r, s[1], { color: ACC, bold: true, fmt: s[2] === "@" ? "General" : s[2] });
  });
  // Row 21: the one input cell (streak target)
  block("B21:D21", "Streak target (edit the yellow cell)", { align: "left", color: MUTED });
  block("E21:G21", target, { bg: "#FFD27A", color: BASE, bold: true, fmt: "0%" });
  home.setRowHeights(15, 7, 28);
  frame("B14:G21");
  home.setRowHeight(22, 14);

  // ---- Segment scores ----
  header("B23:G23", "🏆  SEGMENT SCORES (all time)");
  for (var s = 0; s < nSeg; s++) {
    var row = firstScore + s;
    var wr = 2 + s; // row of this segment on Weekly Progress
    block("B" + row + ":C" + row, "='" + SHEET_WEEKLY + "'!A" + wr, { align: "left" });
    block("D" + row + ":F" + row,
      '=SPARKLINE(G' + row + ',{"charttype","bar";"max",1;"color1","' + ACC + '"})');
    block("G" + row,
      "=IFERROR(AVERAGE(OFFSET('" + SHEET_WEEKLY + "'!$B$" + wr + ",0,0,1,COUNTA('" + SHEET_WEEKLY + "'!$1:$1)-1)),0)",
      { color: ACC, bold: true, fmt: "0%" });
  }
  home.setRowHeights(firstScore, nSeg, 28);
  frame("B23:G" + lastRow);

  // First tab, and the one people land on
  refreshQuote_();
  ss.setActiveSheet(home);
  ss.moveActiveSheet(1);
}
