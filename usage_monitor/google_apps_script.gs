/**
 * Google Sheets backup endpoint for usage_monitor/github_usage_monitor.py.
 *
 * Setup:
 * 1. Create a Google Sheet.
 * 2. Extensions > Apps Script.
 * 3. Paste this file.
 * 4. Optional: set Script property WEBHOOK_SECRET.
 * 5. Deploy > New deployment > Web app > Execute as me > Anyone with the link.
 */

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || "{}");
    const expectedSecret = PropertiesService.getScriptProperties().getProperty("WEBHOOK_SECRET");

    if (expectedSecret && payload.secret !== expectedSecret) {
      return jsonResponse({ ok: false, error: "Invalid secret" });
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    writeSummary(spreadsheet, payload.summary || {});
    writeDaily(spreadsheet, payload.daily || []);
    writeSnapshots(spreadsheet, payload.snapshots || []);
    writeReferrers(spreadsheet, payload.referrers || []);
    writePaths(spreadsheet, payload.paths || []);
    writeReleaseDownloads(spreadsheet, ((payload.releases || {}).assets || []));
    writeDashboard(spreadsheet, payload.summary || {});

    return jsonResponse({
      ok: true,
      updated_at: new Date().toISOString(),
      daily_rows: (payload.daily || []).length,
    });
  } catch (error) {
    return jsonResponse({
      ok: false,
      error: String(error && error.message ? error.message : error),
      stack: String(error && error.stack ? error.stack : ""),
    });
  }
}

function jsonResponse(data, statusCode) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheetByName(spreadsheet, name) {
  return spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
}

function replaceSheet(spreadsheet, name, headers, rows) {
  const sheet = sheetByName(spreadsheet, name);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#d9eaf7");

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

function formatSheet(sheet) {
  sheet.getDataRange()
    .setFontFamily("Arial")
    .setFontSize(10)
    .setVerticalAlignment("middle");
}

function mergeSheetByKey(spreadsheet, name, headers, rows, keyColumnIndex) {
  const sheet = sheetByName(spreadsheet, name);
  const existing = [];

  if (sheet.getLastRow() > 1) {
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
    values.forEach(row => {
      if (row[keyColumnIndex]) {
        existing.push(row);
      }
    });
  }

  const merged = {};
  existing.concat(rows).forEach(row => {
    const key = row[keyColumnIndex];
    if (key) {
      merged[String(key)] = row;
    }
  });

  const sortedRows = Object.keys(merged)
    .sort()
    .map(key => merged[key]);

  replaceSheet(spreadsheet, name, headers, sortedRows);
}

function writeSummary(spreadsheet, summary) {
  const rows = [
    ["Repository", summary.repo || ""],
    ["Last Updated", summary.updated_at || ""],
    ["Tracked Days", summary.tracked_days || 0],
    ["14D Clones", summary.clones_14d || 0],
    ["14D Unique Cloners", summary.unique_cloners_14d || 0],
    ["14D Views", summary.views_14d || 0],
    ["14D Unique Visitors", summary.unique_visitors_14d || 0],
    ["Release Asset Downloads", summary.release_asset_downloads || 0],
  ];
  replaceSheet(spreadsheet, "Summary", ["Metric", "Value"], rows);
  formatSheet(sheetByName(spreadsheet, "Summary"));
}

function writeDaily(spreadsheet, daily) {
  mergeSheetByKey(
    spreadsheet,
    "Daily History",
    ["Date", "Clones", "Unique Cloners", "Views", "Unique Visitors"],
    daily.map(row => [
      row.date || "",
      row.clones || 0,
      row.unique_cloners || 0,
      row.views || 0,
      row.unique_visitors || 0,
    ]),
    0
  );
  formatSheet(sheetByName(spreadsheet, "Daily History"));
}

function writeSnapshots(spreadsheet, snapshots) {
  mergeSheetByKey(
    spreadsheet,
    "14D Snapshots",
    ["Collected At", "Clones 14D", "Unique Cloners 14D", "Views 14D", "Unique Visitors 14D"],
    snapshots.map(row => [
      row.collected_at || "",
      row.clones_14d || 0,
      row.unique_cloners_14d || 0,
      row.views_14d || 0,
      row.unique_visitors_14d || 0,
    ]),
    0
  );
  formatSheet(sheetByName(spreadsheet, "14D Snapshots"));
}

function writeReferrers(spreadsheet, referrers) {
  replaceSheet(
    spreadsheet,
    "Referrers",
    ["Referrer", "Count", "Uniques"],
    referrers.map(row => [
      row.referrer || "",
      row.count || 0,
      row.uniques || 0,
    ])
  );
  formatSheet(sheetByName(spreadsheet, "Referrers"));
}

function writePaths(spreadsheet, paths) {
  replaceSheet(
    spreadsheet,
    "Popular Paths",
    ["Path", "Title", "Count", "Uniques"],
    paths.map(row => [
      row.path || "",
      row.title || "",
      row.count || 0,
      row.uniques || 0,
    ])
  );
  formatSheet(sheetByName(spreadsheet, "Popular Paths"));
}

function writeReleaseDownloads(spreadsheet, assets) {
  replaceSheet(
    spreadsheet,
    "Release Downloads",
    ["Release", "Asset", "Downloads", "Published At"],
    assets.map(row => [
      row.release || "",
      row.asset || "",
      row.downloads || 0,
      row.published_at || "",
    ])
  );
  formatSheet(sheetByName(spreadsheet, "Release Downloads"));
}

function writeDashboard(spreadsheet, summary) {
  const sheet = sheetByName(spreadsheet, "Dashboard");
  sheet.getCharts().forEach(chart => sheet.removeChart(chart));
  clearDashboardSheet(sheet);
  spreadsheet.setActiveSheet(sheet);
  spreadsheet.moveActiveSheet(1);

  sheet.setHiddenGridlines(true);
  sheet.getRange("A1")
    .setValue("GitHub Usage Dashboard")
    .setFontSize(22)
    .setFontWeight("bold")
    .setFontColor("#1d2430");
  sheet.getRange("A2")
    .setValue("Repository: " + (summary.repo || "") + "    Last updated: " + (summary.updated_at || ""))
    .setFontColor("#657184");

  const kpis = [
    ["14D Clones", summary.clones_14d || 0],
    ["14D Unique Cloners", summary.unique_cloners_14d || 0],
    ["14D Views", summary.views_14d || 0],
    ["14D Unique Visitors", summary.unique_visitors_14d || 0],
    ["Release Downloads", summary.release_asset_downloads || 0],
    ["Tracked Days", summary.tracked_days || 0],
  ];

  kpis.forEach((item, index) => {
    const startCol = 1 + (index % 3) * 3;
    const startRow = 4 + Math.floor(index / 3) * 4;
    const range = sheet.getRange(startRow, startCol, 3, 2);
    range
      .setBackground("#ffffff")
      .setBorder(true, true, true, true, false, false, "#d9dee7", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(startRow, startCol)
      .setValue(item[0])
      .setFontColor("#657184")
      .setFontWeight("bold");
    sheet.getRange(startRow + 1, startCol)
      .setValue(item[1])
      .setNumberFormat("#,##0")
      .setFontSize(24)
      .setFontWeight("bold")
      .setFontColor("#111827")
      .setVerticalAlignment("middle");
  });

  const dailySheet = sheetByName(spreadsheet, "Daily History");
  const referrersSheet = sheetByName(spreadsheet, "Referrers");
  const releasesSheet = sheetByName(spreadsheet, "Release Downloads");
  const pathsSheet = sheetByName(spreadsheet, "Popular Paths");

  addDailyChart(sheet, dailySheet);
  addBarChart(sheet, referrersSheet, 1, 2, "Top Referrers", 27, 1);
  addBarChart(sheet, releasesSheet, 2, 3, "Release Asset Downloads", 27, 5);
  addBarChart(sheet, pathsSheet, 1, 3, "Popular Pages", 44, 1);

  sheet.getRange("A10")
    .setValue("GitHub Traffic API는 최근 14일만 제공하므로, 이 시트는 매일 실행된 백업 결과를 누적해 장기 추세를 보여줍니다.")
    .setFontColor("#657184")
    .setFontSize(9);

  for (let column = 1; column <= 8; column++) {
    sheet.setColumnWidth(column, column % 3 === 0 ? 26 : 150);
  }
  for (let row = 1; row <= 60; row++) {
    sheet.setRowHeight(row, 24);
  }
}

function clearDashboardSheet(sheet) {
  try {
    const mergedRanges = sheet
      .getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
      .getMergedRanges();
    mergedRanges.forEach(range => range.breakApart());
  } catch (error) {
    // Older dashboard versions used merged cells. If unmerge fails, continue
    // with a non-merged layout so the backup itself does not fail.
  }
  sheet.clear();
}

function addDailyChart(dashboardSheet, dailySheet) {
  const lastRow = dailySheet.getLastRow();
  if (lastRow < 2) {
    dashboardSheet.getRange("A12").setValue("일별 추세 데이터가 아직 없습니다.");
    return;
  }

  const chart = dashboardSheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(dailySheet.getRange(1, 1, lastRow, 5))
    .setOption("title", "Daily Clone/View Trend")
    .setOption("legend", { position: "bottom" })
    .setOption("curveType", "function")
    .setOption("height", 330)
    .setOption("width", 930)
    .setOption("chartArea", { left: 60, top: 45, width: "82%", height: "68%" })
    .setOption("colors", ["#2563eb", "#0f8b6f", "#b7791f", "#c2415d"])
    .setPosition(12, 1, 0, 0)
    .build();
  dashboardSheet.insertChart(chart);
}

function addBarChart(dashboardSheet, sourceSheet, labelColumn, valueColumn, title, row, column) {
  const lastRow = Math.min(sourceSheet.getLastRow(), 9);
  if (lastRow < 2) {
    dashboardSheet.getRange(row, column).setValue(title + " 데이터가 아직 없습니다.");
    return;
  }

  const helperStartColumn = 10 + column;
  const helperRange = dashboardSheet.getRange(row, helperStartColumn, lastRow - 1, 2);
  const values = sourceSheet.getRange(2, 1, lastRow - 1, sourceSheet.getLastColumn()).getValues()
    .map(sourceRow => [sourceRow[labelColumn - 1], sourceRow[valueColumn - 1]]);
  helperRange.setValues(values);

  const chart = dashboardSheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(helperRange)
    .setOption("title", title)
    .setOption("legend", { position: "none" })
    .setOption("height", 300)
    .setOption("width", 430)
    .setOption("chartArea", { left: 120, top: 45, width: "62%", height: "70%" })
    .setOption("colors", ["#2563eb"])
    .setPosition(row, column, 0, 0)
    .build();
  dashboardSheet.insertChart(chart);
  helperRange.setFontColor("#ffffff");
}
