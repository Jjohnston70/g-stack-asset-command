// ============================================
// G-STACK ASSET COMMAND FREE STARTER KIT
// True North Data Strategies
// Version 1.0
// ============================================

/**
 * Creates the Tools menu when the spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🚗 Tools")
    .addItem("🖥️ Open Dashboard", "showDashboard")
    .addItem("🌐 Open Dashboard in Browser", "openDashboardInBrowser")
    .addItem(
      "🔐 Allow This Sheet for Web View",
      "addCurrentSpreadsheetToAllowList",
    )
    .addItem("📝 Data Entry", "openDataEntrySidebar")
    .addItem("📚 User Manual", "showUserManual")
    .addSeparator()
    .addItem("1. Build Complete Template", "buildAssetCommandTemplate")
    .addItem("2. Setup Dashboard 2", "setupDashboard2")
    .addItem("3. Add Test Data", "addTestData")
    .addSeparator()
    .addItem("Clear Test Data", "clearTestData")
    .addItem("Refresh Dashboards", "refreshDashboards")
    .addSeparator()
    .addItem("Send Daily Digest", "sendDailyDigest")
    .addItem("Check Maintenance Due", "checkMaintenanceDue")
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("🚛 Driver Compliance")
        .addItem("Check Driver Credentials", "checkDriverCompliance")
        .addItem("Check Fuel Anomalies", "checkFuelAnomaly")
        .addSeparator()
        .addItem("Add Driver Test Data", "addDriverTestData")
        .addItem("Clear Driver Data", "clearDriverTestData"),
    )
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("⚙️ Function Runner")
        .addItem("Initialize Function Runner", "initializeFunctionRunnerSheet")
        .addSeparator()
        .addItem("Setup Trigger", "setupFunctionRunnerTrigger")
        .addItem("Remove Trigger", "removeFunctionRunnerTrigger"),
    )
    .addToUi();
}
function doGet(e) {
  try {
    const spreadsheetId = e && e.parameter ? e.parameter.sid : "";
    const ss = getAssetCommandSpreadsheet_(spreadsheetId);
    const template = HtmlService.createTemplateFromFile("Dashboard");
    template.spreadsheetId = ss.getId();
    return template.evaluate().setTitle("G-Stack Asset Command Dashboard");
  } catch (error) {
    return HtmlService.createHtmlOutput(
      "<h3>G-Stack Asset Command Setup Required</h3>" + "<p>" + error.message + "</p>",
    );
  }
}

/**
 * Run this ONCE from the Script Editor after each new deployment.
 * Paste your /exec URL inside the quotes below.
 */
function setWebAppUrl() {
  PropertiesService.getScriptProperties().setProperty(
    "ASSETCOMMAND_WEBAPP_URL",
    "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec",
  );
  SpreadsheetApp.getUi().alert("Web app URL saved.");
}

function openDashboardInBrowser() {
  const url = PropertiesService.getScriptProperties().getProperty(
    "ASSETCOMMAND_WEBAPP_URL",
  );

  if (!url) {
    SpreadsheetApp.getUi().alert(
      "Run setWebAppUrl() from Script Editor first.",
    );
    return;
  }

  const html = HtmlService.createHtmlOutput(
    "<!DOCTYPE html><html><body>" +
      '<p style="font-family:Arial;font-size:14px;">Click below to open your dashboard:</p>' +
      '<input type="button" value="🚀 Open Dashboard" ' +
      'style="background:#1a3a5c;color:white;border:none;padding:12px 24px;font-size:15px;border-radius:6px;cursor:pointer;" ' +
      "onclick=\"window.open('" +
      url +
      "','_blank');\" />" +
      "</body></html>",
  )
    .setWidth(400)
    .setHeight(120);

  SpreadsheetApp.getUi().showModalDialog(html, "G-Stack Asset Command Dashboard");
}

/**
 * Adds the current spreadsheet ID to the allowed web-view list.
 */
function addCurrentSpreadsheetToAllowList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("No active spreadsheet context.");
  const allowed = getAllowedSpreadsheetIds_();
  if (!allowed.includes(ss.getId())) {
    allowed.push(ss.getId());
    saveAllowedSpreadsheetIds_(allowed);
  }
  SpreadsheetApp.getUi().alert(
    "This spreadsheet is now allowed for web dashboard viewing.",
  );
}

/**
 * Sets allowed spreadsheet IDs from a comma-separated string.
 * Example: setAllowedSpreadsheetIds('id1,id2,id3')
 */
function setAllowedSpreadsheetIds(idsCsv) {
  const ids = (idsCsv || "")
    .toString()
    .split(",")
    .map((id) => id.trim())
    .filter(Boolean);
  saveAllowedSpreadsheetIds_(ids);
}

function getAllowedSpreadsheetIds_() {
  const props = PropertiesService.getScriptProperties();
  const raw = (
    props.getProperty("ASSETCOMMAND_ALLOWED_SPREADSHEET_IDS") || ""
  ).trim();
  return raw
    ? raw
        .split(",")
        .map((id) => id.trim())
        .filter(Boolean)
    : [];
}

function saveAllowedSpreadsheetIds_(ids) {
  const unique = Array.from(
    new Set((ids || []).map((id) => id.trim()).filter(Boolean)),
  );
  PropertiesService.getScriptProperties().setProperty(
    "ASSETCOMMAND_ALLOWED_SPREADSHEET_IDS",
    unique.join(","),
  );
}

/**
 * Resolves the spreadsheet in both bound-sheet and web-app contexts.
 */
function getAssetCommandSpreadsheet_(spreadsheetId) {
  const requestedId = (spreadsheetId || "").toString().trim();
  const allowedIds = getAllowedSpreadsheetIds_();

  if (requestedId) {
    if (!allowedIds.includes(requestedId)) {
      throw new Error("Invalid spreadsheet context.");
    }
    return SpreadsheetApp.openById(requestedId);
  }

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    const activeId = active.getId();
    if (!allowedIds.includes(activeId)) {
      allowedIds.push(activeId);
      saveAllowedSpreadsheetIds_(allowedIds);
    }
    return active;
  }

  throw new Error(
    "No spreadsheet context found. Open the sheet and run Allow This Sheet for Web View from the menu.",
  );
}

/**
 * Refreshes all dashboard formulas
 */
function refreshDashboards() {
  SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Dashboard")
    ?.getRange("A1")
    .getValue();
  SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Dashboard 2")
    ?.getRange("A1")
    .getValue();
  SpreadsheetApp.getUi().alert("Dashboards refreshed!");
}

function fixAllowList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const id = ss.getId();
  const allowed = getAllowedSpreadsheetIds_();
  if (!allowed.includes(id)) {
    allowed.push(id);
    saveAllowedSpreadsheetIds_(allowed);
  }
  PropertiesService.getScriptProperties().setProperty(
    "ASSETCOMMAND_SPREADSHEET_ID",
    id,
  );
  SpreadsheetApp.getUi().alert("Allow list updated. ID: " + id);
}

// ============================================
// TEMPLATE BUILDER
// ============================================

/**
 * Builds the complete AssetCommand template with all sheets
 */
function buildAssetCommandTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Check if template already exists
  if (ss.getSheetByName("Assets Master")) {
    const response = ui.alert(
      "Template Already Exists",
      "This will delete existing sheets and rebuild. Continue?",
      ui.ButtonSet.YES_NO,
    );
    if (response !== ui.Button.YES) return;

    // Delete existing sheets
    const sheetsToDelete = [
      "Dashboard",
      "Assets Master",
      "Drivers",
      "Activity Log",
      "Maintenance Tracker",
      "Cost Tracking",
      "Config",
      "Setup Instructions",
      "Dashboard 2",
    ];
    sheetsToDelete.forEach((name) => {
      const sheet = ss.getSheetByName(name);
      if (sheet) ss.deleteSheet(sheet);
    });
  }

  // Create temporary sheet if needed
  let tempSheet = ss.getSheetByName("Sheet1") || ss.insertSheet("Temp");
  if (tempSheet.getName() !== "Temp") {
    tempSheet.setName("Temp");
  }

  // === CREATE ALL SHEETS ===
  const dashboard = ss.insertSheet("Dashboard");
  const assets = ss.insertSheet("Assets Master");
  const drivers = ss.insertSheet("Drivers");
  const activity = ss.insertSheet("Activity Log");
  const maint = ss.insertSheet("Maintenance Tracker");
  const cost = ss.insertSheet("Cost Tracking");
  const config = ss.insertSheet("Config");
  const setup = ss.insertSheet("Setup Instructions");

  // ============================================
  // ASSETS MASTER SHEET
  // ============================================
  const assetHeaders = [
    "Asset ID",
    "Asset Name",
    "Asset Type",
    "Status",
    "Current Location",
    "Assigned To",
    "Last Service Date",
    "Next Service Due",
    "Service Interval (Days)",
    "Acquisition Date",
    "Acquisition Cost",
    "Notes",
  ];
  assets
    .getRange(1, 1, 1, assetHeaders.length)
    .setValues([assetHeaders])
    .setFontWeight("bold")
    .setBackground("#1E3A5F")
    .setFontColor("white");

  // Data validation dropdowns
  const assetTypes = ["Vehicle", "Equipment", "Tool", "Trailer", "Other"];
  const statuses = ["Available", "In Use", "Maintenance", "Retired"];

  assets
    .getRange("C2:C100")
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(assetTypes, true)
        .build(),
    );
  assets
    .getRange("D2:D100")
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(statuses, true)
        .build(),
    );

  // Auto-resize columns
  assets.autoResizeColumns(1, assetHeaders.length);

  // Freeze header row
  assets.setFrozenRows(1);

  // ============================================
  // DRIVERS SHEET (DOT Compliance)
  // ============================================
  const driverHeaders = [
    // Identity (cols A-E)
    "Driver ID",
    "Driver Name",
    "Status",
    "Phone",
    "Email",
    // License (cols F-I)
    "License Number",
    "License State",
    "License Expiry",
    "License Type",
    // CDL (cols J-L)
    "CDL Number",
    "CDL Expiry",
    "CDL Endorsements",
    // Medical (cols M-O)
    "Medical Card Expiry",
    "Medical Examiner",
    "Medical Exam Date",
    // Background (cols P-R)
    "Background Check Date",
    "Clearinghouse Status",
    "MVR Review Date",
    // Drug Testing (cols S-T)
    "Drug Test Date",
    "Drug Test Result",
    // HOS Compliance (cols U-X)
    "Last HOS Audit",
    "ELD Provider",
    "ELD Device ID",
    "ELD Compliant",
    // Vehicle Inspections - DVIR (cols Y-AA)
    "Last Pre-Trip Date",
    "Last Post-Trip Date",
    "Annual Inspection Due",
    // Employment (cols AB-AD)
    "Hire Date",
    "Termination Date",
    "Notes",
  ];
  drivers
    .getRange(1, 1, 1, driverHeaders.length)
    .setValues([driverHeaders])
    .setFontWeight("bold")
    .setBackground("#1E3A5F")
    .setFontColor("white");

  // Data validation dropdowns for Drivers
  const driverStatuses = ["Active", "On Leave", "Suspended", "Terminated"];
  const licenseTypes = ["Class A", "Class B", "Class C", "Non-CDL"];
  const testResults = ["Negative", "Positive", "Pending", "Refused"];
  const clearinghouseStatus = ["Clear", "Violation", "Pending Query"];
  const yesNo = ["Yes", "No", "Pending"];
  const eldProviders = [
    "KeepTruckin",
    "Samsara",
    "Omnitracs",
    "PeopleNet",
    "Geotab",
    "Other",
  ];

  // Status dropdown (col C)
  drivers
    .getRange("C2:C100")
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(driverStatuses, true)
        .build(),
    );
  // License Type dropdown (col I)
  drivers
    .getRange("I2:I100")
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(licenseTypes, true)
        .build(),
    );
  // Clearinghouse Status dropdown (col R)
  drivers
    .getRange("R2:R100")
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(clearinghouseStatus, true)
        .build(),
    );
  // Drug Test Result dropdown (col T)
  drivers
    .getRange("T2:T100")
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(testResults, true)
        .build(),
    );
  // ELD Provider dropdown (col V)
  drivers
    .getRange("V2:V100")
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(eldProviders, true)
        .build(),
    );
  // ELD Compliant dropdown (col X)
  drivers
    .getRange("X2:X100")
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(yesNo, true)
        .build(),
    );

  // Date columns formatting (H=8, K=11, M=13, O=15, P=16, R=18, S=19, U=21, Y=25, Z=26, AA=27, AB=28, AC=29)
  const dateColumns = [8, 11, 13, 15, 16, 18, 19, 21, 25, 26, 27, 28, 29];
  dateColumns.forEach((col) => {
    drivers.getRange(2, col, 99, 1).setNumberFormat("yyyy-mm-dd");
  });

  drivers.autoResizeColumns(1, driverHeaders.length);
  drivers.setFrozenRows(1);

  // ============================================
  // ACTIVITY LOG SHEET
  // ============================================
  const actHeaders = [
    "Timestamp",
    "Asset ID",
    "Asset Name",
    "Action Type",
    "Employee",
    "Location",
    "Odometer/Hours",
    "Fuel Added (gal)",
    "Fuel Cost",
    "Condition Notes",
  ];
  activity
    .getRange(1, 1, 1, actHeaders.length)
    .setValues([actHeaders])
    .setFontWeight("bold")
    .setBackground("#1E3A5F")
    .setFontColor("white");

  const actions = [
    "Check-out",
    "Check-in",
    "Refuel",
    "Maintenance",
    "Incident",
    "Location Update",
  ];
  activity
    .getRange("D2:D500")
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(actions, true)
        .build(),
    );

  activity.autoResizeColumns(1, actHeaders.length);
  activity.setFrozenRows(1);

  // ============================================
  // MAINTENANCE TRACKER SHEET
  // ============================================
  const maintHeaders = [
    "Maintenance ID",
    "Asset ID",
    "Asset Name",
    "Service Type",
    "Scheduled Date",
    "Completed Date",
    "Status",
    "Parts Cost",
    "Labor Cost",
    "Total Cost",
    "Vendor",
    "Invoice Number",
    "Notes",
  ];
  maint
    .getRange(1, 1, 1, maintHeaders.length)
    .setValues([maintHeaders])
    .setFontWeight("bold")
    .setBackground("#1E3A5F")
    .setFontColor("white");

  const services = [
    "Oil Change",
    "Tire Rotation",
    "Inspection",
    "Repair",
    "Blade Sharpening",
    "Filter Replacement",
    "Annual Inspection",
    "DOT Physical",
    "Other",
  ];
  maint
    .getRange("D2:D200")
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(services, true)
        .build(),
    );

  maint.autoResizeColumns(1, maintHeaders.length);
  maint.setFrozenRows(1);

  // ============================================
  // COST TRACKING SHEET
  // ============================================
  const costHeaders = [
    "Asset ID",
    "Asset Name",
    "Total Fuel Cost (30 days)",
    "Total Maintenance Cost (YTD)",
    "Number of Trips",
    "Avg Cost Per Trip",
    "Days Since Last Service",
  ];
  cost
    .getRange(1, 1, 1, costHeaders.length)
    .setValues([costHeaders])
    .setFontWeight("bold")
    .setBackground("#1E3A5F")
    .setFontColor("white");

  cost.autoResizeColumns(1, costHeaders.length);
  cost.setFrozenRows(1);

  // ============================================
  // CONFIG SHEET
  // ============================================
  const configData = [
    ["Setting", "Value"],
    ["Business Name", ""],
    ["Email Address", ""],
    ["Alert Threshold (days)", 7],
    ["Fuel Anomaly Threshold (%)", 20],
  ];
  config.getRange(1, 1, configData.length, 2).setValues(configData);
  config
    .getRange(1, 1, 1, 2)
    .setFontWeight("bold")
    .setBackground("#1E3A5F")
    .setFontColor("white");
  config.getRange("A2:A5").setFontWeight("bold");
  config.autoResizeColumns(1, 2);

  // ============================================
  // SETUP INSTRUCTIONS SHEET
  // ============================================
  const setupContent = [
    ["🚀 G-STACK ASSET COMMAND FREE STARTER KIT - SETUP GUIDE"],
    [""],
    ["GETTING STARTED:"],
    [""],
    ["Step 1: You just ran 'Build Complete Template' ✓"],
    ["Step 2: Run 🚗 Tools → Setup Dashboard 2"],
    ["Step 3: Run 🚗 Tools → Add Test Data (to see it working)"],
    ["Step 4: Go to Config sheet and enter your business name & email"],
    ["Step 5: Clear test data and add your own assets!"],
    [""],
    ["━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"],
    [""],
    ["MENU REFERENCE (🚗 Tools menu in toolbar):"],
    [""],
    ["• Build Complete Template - Creates all sheets (first time only)"],
    ["• Setup Dashboard 2 - Creates executive dashboard"],
    ["• Add Test Data - Adds sample data to test formulas"],
    ["• Clear Test Data - Removes sample data, keeps structure"],
    ["• Refresh Dashboards - Recalculates all formulas"],
    ["• Send Daily Digest - Emails asset summary"],
    ["• Check Maintenance Due - Emails maintenance alerts"],
    [""],
    ["━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"],
    [""],
    ["COLOR GUIDE:"],
    ["🟢 Green = Good (Available, current)"],
    ["🟡 Yellow/Orange = Attention (In Use, due soon)"],
    ["🔴 Red = Action Required (Overdue, retired)"],
    [""],
    ["━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"],
    [""],
    ["NEED HELP?"],
    ["Email: jacob@truenorthstrategyops.com"],
    ["Web: truenorthstrategyops.com/solutions/asset-command"],
    [""],
    ["━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"],
    [""],
    ["© 2025 True North Data Strategies LLC"],
  ];
  setup.getRange(1, 1, setupContent.length, 1).setValues(setupContent);
  setup
    .getRange("A1")
    .setFontSize(16)
    .setFontWeight("bold")
    .setFontColor("#00B4D8");
  setup.getRange("A3").setFontWeight("bold");
  setup.getRange("A13").setFontWeight("bold");
  setup.getRange("A24").setFontWeight("bold");
  setup.getRange("A31").setFontWeight("bold");
  setup.setColumnWidth(1, 500);

  // ============================================
  // DASHBOARD SHEET
  // ============================================
  dashboard
    .getRange("A1")
    .setValue("G-STACK ASSET COMMAND - FREE STARTER KIT")
    .setFontSize(22)
    .setFontWeight("bold")
    .setFontColor("#00B4D8");
  dashboard
    .getRange("A2")
    .setFormula('="Last Updated: "&TEXT(NOW(),"mmmm d, yyyy h:mm AM/PM")')
    .setFontStyle("italic")
    .setFontColor("#666666");

  // KPI Section
  dashboard
    .getRange("A4")
    .setValue("📊 QUICK STATS")
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground("#E2E8F0");

  // KPI Labels
  dashboard.getRange("A5").setValue("Total Assets");
  dashboard.getRange("A6").setValue("Available");
  dashboard.getRange("A7").setValue("In Use");
  dashboard.getRange("A8").setValue("Maintenance");
  dashboard.getRange("A9").setValue("Retired");
  dashboard.getRange("A5:A9").setFontWeight("bold");

  // KPI Formulas
  dashboard.getRange("B5").setFormula("=COUNTA('Assets Master'!A2:A)");
  dashboard
    .getRange("B6")
    .setFormula("=COUNTIF('Assets Master'!D2:D,\"Available\")");
  dashboard
    .getRange("B7")
    .setFormula("=COUNTIF('Assets Master'!D2:D,\"In Use\")");
  dashboard
    .getRange("B8")
    .setFormula("=COUNTIF('Assets Master'!D2:D,\"Maintenance\")");
  dashboard
    .getRange("B9")
    .setFormula("=COUNTIF('Assets Master'!D2:D,\"Retired\")");

  // KPI Formatting
  dashboard
    .getRange("B5:B9")
    .setFontWeight("bold")
    .setFontSize(16)
    .setHorizontalAlignment("center");
  dashboard.getRange("B5").setBackground("#3B82F6").setFontColor("white"); // Total - Blue
  dashboard.getRange("B6").setBackground("#10B981").setFontColor("white"); // Available - Green
  dashboard.getRange("B7").setBackground("#F59E0B").setFontColor("white"); // In Use - Orange
  dashboard.getRange("B8").setBackground("#EF4444").setFontColor("white"); // Maintenance - Red
  dashboard.getRange("B9").setBackground("#6B7280").setFontColor("white"); // Retired - Gray

  // Status Summary for Chart
  dashboard
    .getRange("D4")
    .setValue("📈 STATUS BREAKDOWN")
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground("#E2E8F0");
  dashboard
    .getRange("D5:E5")
    .setValues([["Status", "Count"]])
    .setFontWeight("bold")
    .setBackground("#CBD5E1");
  dashboard.getRange("D6").setValue("Available");
  dashboard.getRange("D7").setValue("In Use");
  dashboard.getRange("D8").setValue("Maintenance");
  dashboard.getRange("D9").setValue("Retired");
  dashboard
    .getRange("E6")
    .setFormula("=COUNTIF('Assets Master'!D2:D,\"Available\")");
  dashboard
    .getRange("E7")
    .setFormula("=COUNTIF('Assets Master'!D2:D,\"In Use\")");
  dashboard
    .getRange("E8")
    .setFormula("=COUNTIF('Assets Master'!D2:D,\"Maintenance\")");
  dashboard
    .getRange("E9")
    .setFormula("=COUNTIF('Assets Master'!D2:D,\"Retired\")");

  // Create Pie Chart
  const chart = dashboard
    .newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(dashboard.getRange("D5:E9"))
    .setPosition(11, 1, 0, 0)
    .setOption("title", "Asset Status Distribution")
    .setOption("pieSliceText", "percentage")
    .setOption("legend", { position: "right" })
    .setOption("colors", ["#10B981", "#F59E0B", "#EF4444", "#6B7280"])
    .setOption("width", 450)
    .setOption("height", 280)
    .build();
  dashboard.insertChart(chart);

  // Maintenance Alerts Section
  dashboard
    .getRange("A24")
    .setValue("⚠️ MAINTENANCE ALERTS")
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground("#FEE2E2");
  dashboard
    .getRange("A25")
    .setValue("Assets with maintenance due in next 7 days:")
    .setFontStyle("italic");
  dashboard
    .getRange("A26")
    .setFormula(
      '=IFERROR(QUERY(\'Assets Master\'!A2:H,"SELECT A,B,H WHERE H <= date \'"&TEXT(TODAY()+7,"yyyy-mm-dd")&"\' AND H > date \'"&TEXT(TODAY(),"yyyy-mm-dd")&"\' ORDER BY H LIMIT 5"),"No upcoming maintenance")',
    );

  dashboard.autoResizeColumns(1, 5);

  // Delete temporary sheet
  const temp = ss.getSheetByName("Temp");
  if (temp) ss.deleteSheet(temp);

  ui.alert(
    "✅ Template Built Successfully!\n\n" +
      "Next Steps:\n" +
      "1. Run: 🚗 Tools → Setup Dashboard 2\n" +
      "2. Run: 🚗 Tools → Add Test Data\n" +
      "3. Check both dashboards!",
  );

  Logger.log("AssetCommand template built successfully");
}

// ============================================
// DASHBOARD 2 (EXECUTIVE VIEW)
// ============================================

/**
 * Creates Dashboard 2 with executive summary view
 */
function setupDashboard2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Check if Assets Master exists
  if (!ss.getSheetByName("Assets Master")) {
    ui.alert('Please run "Build Complete Template" first.');
    return;
  }

  // Delete existing Dashboard 2 if present
  const existing = ss.getSheetByName("Dashboard 2");
  if (existing) ss.deleteSheet(existing);

  const dash = ss.insertSheet("Dashboard 2");

  // === HEADER ===
  dash
    .getRange("A1")
    .setValue("G-Stack Asset Command Free Starter Kit")
    .setFontSize(20)
    .setFontWeight("bold")
    .setFontColor("white")
    .setBackground("#0F172A");
  dash.getRange("B1:H1").setBackground("#0F172A");

  dash
    .getRange("A2")
    .setFormula(
      '=IF(Config!B2<>"",Config!B2&" | Executive Summary","Executive Summary")',
    )
    .setFontSize(12)
    .setBackground("#E2E8F0");
  dash
    .getRange("A3")
    .setFormula('=TEXT(TODAY(),"dddd, mmmm d, yyyy")')
    .setFontStyle("italic")
    .setBackground("#E2E8F0");

  // === KPI ROW ===
  dash
    .getRange("A5")
    .setValue("📊 KEY METRICS")
    .setFontWeight("bold")
    .setFontSize(12);

  // KPI Boxes
  const kpiLabels = [
    "Total Assets",
    "Available",
    "In Use",
    "Maint. Due",
    "Overdue",
  ];
  dash
    .getRange("A6:E6")
    .setValues([kpiLabels])
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#CBD5E1");

  dash.getRange("A7").setFormula("=COUNTA('Assets Master'!A2:A)");
  dash
    .getRange("B7")
    .setFormula("=COUNTIF('Assets Master'!D2:D,\"Available\")");
  dash.getRange("C7").setFormula("=COUNTIF('Assets Master'!D2:D,\"In Use\")");
  dash
    .getRange("D7")
    .setFormula(
      "=COUNTIFS('Assets Master'!H2:H,\"<=\"&TODAY()+7,'Assets Master'!H2:H,\">\"&TODAY())",
    );
  dash
    .getRange("E7")
    .setFormula("=COUNTIF('Maintenance Tracker'!G2:G,\"Overdue\")");

  dash
    .getRange("A7:E7")
    .setFontWeight("bold")
    .setFontSize(18)
    .setHorizontalAlignment("center");
  dash.getRange("A7").setBackground("#3B82F6").setFontColor("white");
  dash.getRange("B7").setBackground("#10B981").setFontColor("white");
  dash.getRange("C7").setBackground("#F59E0B").setFontColor("white");
  dash.getRange("D7").setBackground("#F97316").setFontColor("white");
  dash.getRange("E7").setBackground("#EF4444").setFontColor("white");

  // === RECENT ACTIVITY ===
  dash
    .getRange("A9")
    .setValue("🕒 RECENT ACTIVITY")
    .setFontWeight("bold")
    .setFontSize(12);
  dash
    .getRange("A10:D10")
    .setValues([["Date", "Asset", "Name", "Action"]])
    .setFontWeight("bold")
    .setBackground("#CBD5E1");

  dash
    .getRange("A11")
    .setFormula(
      '=IFERROR(QUERY(\'Activity Log\'!A2:D,"SELECT A,B,C,D WHERE A IS NOT NULL ORDER BY A DESC LIMIT 8"),"No activity logged yet")',
    );

  // === STATUS SUMMARY TABLE FOR CHART ===
  dash
    .getRange("F9")
    .setValue("📈 STATUS BREAKDOWN")
    .setFontWeight("bold")
    .setFontSize(12);
  dash
    .getRange("F10:G10")
    .setValues([["Status", "Count"]])
    .setFontWeight("bold")
    .setBackground("#CBD5E1");
  dash.getRange("F11").setValue("Available");
  dash.getRange("F12").setValue("In Use");
  dash.getRange("F13").setValue("Maintenance");
  dash.getRange("F14").setValue("Retired");
  dash
    .getRange("G11")
    .setFormula("=COUNTIF('Assets Master'!D2:D,\"Available\")");
  dash.getRange("G12").setFormula("=COUNTIF('Assets Master'!D2:D,\"In Use\")");
  dash
    .getRange("G13")
    .setFormula("=COUNTIF('Assets Master'!D2:D,\"Maintenance\")");
  dash.getRange("G14").setFormula("=COUNTIF('Assets Master'!D2:D,\"Retired\")");

  // === PIE CHART ===
  const chart = dash
    .newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(dash.getRange("F10:G14"))
    .setPosition(16, 6, 0, 0)
    .setOption("title", "Asset Status")
    .setOption("pieSliceText", "value")
    .setOption("legend", { position: "labeled" })
    .setOption("colors", ["#10B981", "#F59E0B", "#EF4444", "#6B7280"])
    .setOption("width", 350)
    .setOption("height", 220)
    .build();
  dash.insertChart(chart);

  // === MAINTENANCE STATUS TABLE FOR CHART ===
  dash
    .getRange("F24")
    .setValue("🔧 MAINTENANCE STATUS")
    .setFontWeight("bold")
    .setFontSize(12);
  dash
    .getRange("F25:G25")
    .setValues([["Status", "Count"]])
    .setFontWeight("bold")
    .setBackground("#CBD5E1");
  dash.getRange("F26").setValue("Scheduled");
  dash.getRange("F27").setValue("Completed");
  dash.getRange("F28").setValue("Overdue");
  dash
    .getRange("G26")
    .setFormula("=COUNTIF('Maintenance Tracker'!G2:G,\"Scheduled\")");
  dash
    .getRange("G27")
    .setFormula("=COUNTIF('Maintenance Tracker'!G2:G,\"Completed\")");
  dash
    .getRange("G28")
    .setFormula("=COUNTIF('Maintenance Tracker'!G2:G,\"Overdue\")");

  // === COLUMN CHART FOR MAINTENANCE ===
  const chart2 = dash
    .newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dash.getRange("F25:G28"))
    .setPosition(30, 6, 0, 0)
    .setOption("title", "Maintenance Overview")
    .setOption("legend", { position: "none" })
    .setOption("colors", ["#3B82F6"])
    .setOption("width", 350)
    .setOption("height", 200)
    .build();
  dash.insertChart(chart2);

  // Auto-resize
  dash.autoResizeColumns(1, 7);

  ui.alert(
    "✅ Dashboard 2 Created!\n\n" +
      'Run "Add Test Data" to see charts populate.',
  );

  Logger.log("Dashboard 2 created successfully");
}

// ============================================
// TEST DATA FUNCTIONS
// ============================================

/**
 * Adds sample test data to all sheets
 */
function addTestData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Verify template exists
  if (!ss.getSheetByName("Assets Master")) {
    ui.alert('Please run "Build Complete Template" first.');
    return;
  }

  const today = new Date();

  // ============================================
  // ASSETS MASTER DATA
  // ============================================
  const assets = ss.getSheetByName("Assets Master");

  const assetData = [
    [
      "A-001",
      "Red F-150",
      "Vehicle",
      "Available",
      "Main Yard",
      "John Smith",
      new Date(2024, 10, 1),
      null,
      90,
      new Date(2023, 5, 15),
      35000,
      "Fleet truck #1",
    ],
    [
      "A-002",
      "Blue Silverado",
      "Vehicle",
      "In Use",
      "Job Site A",
      "Mike Johnson",
      new Date(2024, 9, 15),
      null,
      90,
      new Date(2023, 8, 20),
      38000,
      "Fleet truck #2",
    ],
    [
      "A-003",
      "CAT Excavator",
      "Equipment",
      "Available",
      "Main Yard",
      "",
      new Date(2024, 11, 1),
      null,
      180,
      new Date(2022, 3, 10),
      125000,
      "Mini excavator",
    ],
    [
      "A-004",
      "John Deere Mower",
      "Equipment",
      "Maintenance",
      "Shop",
      "Service Dept",
      new Date(2024, 8, 1),
      null,
      30,
      new Date(2023, 4, 1),
      8500,
      "Commercial mower",
    ],
    [
      "A-005",
      "Utility Trailer",
      "Trailer",
      "Available",
      "Main Yard",
      "",
      new Date(2024, 6, 15),
      null,
      365,
      new Date(2021, 1, 1),
      4500,
      "16ft flatbed",
    ],
    [
      "A-006",
      "Pressure Washer",
      "Equipment",
      "In Use",
      "Job Site B",
      "Sarah Wilson",
      new Date(2024, 11, 10),
      null,
      60,
      new Date(2024, 2, 1),
      2200,
      "Commercial grade",
    ],
    [
      "A-007",
      "White Transit Van",
      "Vehicle",
      "Available",
      "Main Yard",
      "",
      new Date(2024, 10, 20),
      null,
      90,
      new Date(2024, 1, 15),
      42000,
      "Cargo van",
    ],
    [
      "A-008",
      "Generator 5000W",
      "Equipment",
      "Retired",
      "Storage",
      "",
      new Date(2023, 5, 1),
      null,
      180,
      new Date(2019, 6, 1),
      1500,
      "Needs replacement",
    ],
    [
      "A-009",
      "Bobcat Skid Steer",
      "Equipment",
      "In Use",
      "Job Site C",
      "Tom Brown",
      new Date(2024, 11, 5),
      null,
      120,
      new Date(2022, 9, 1),
      45000,
      "S650 model",
    ],
    [
      "A-010",
      "Enclosed Trailer",
      "Trailer",
      "Available",
      "Main Yard",
      "",
      new Date(2024, 7, 1),
      null,
      365,
      new Date(2023, 3, 15),
      8000,
      "20ft enclosed",
    ],
  ];

  assets.getRange(2, 1, assetData.length, 12).setValues(assetData);

  // Add Next Service Due formula to each row
  for (let i = 2; i <= assetData.length + 1; i++) {
    assets
      .getRange(i, 8)
      .setFormula(
        "=IF(OR(ISBLANK(G" +
          i +
          "),ISBLANK(I" +
          i +
          ')),"",G' +
          i +
          "+I" +
          i +
          ")",
      );
  }

  // ============================================
  // ACTIVITY LOG DATA
  // ============================================
  const activity = ss.getSheetByName("Activity Log");

  const activityData = [
    [
      new Date(today.getTime() - 1 * 24 * 60 * 60 * 1000),
      "A-001",
      null,
      "Check-out",
      "John Smith",
      "Main Yard",
      45230,
      null,
      null,
      "Morning pickup",
    ],
    [
      new Date(today.getTime() - 1 * 24 * 60 * 60 * 1000),
      "A-001",
      null,
      "Refuel",
      "John Smith",
      "Shell Station",
      45280,
      18.5,
      62.5,
      "Regular unleaded",
    ],
    [
      new Date(today.getTime() - 1 * 24 * 60 * 60 * 1000),
      "A-001",
      null,
      "Check-in",
      "John Smith",
      "Main Yard",
      45320,
      null,
      null,
      "End of day",
    ],
    [
      new Date(today.getTime() - 2 * 24 * 60 * 60 * 1000),
      "A-002",
      null,
      "Check-out",
      "Mike Johnson",
      "Main Yard",
      38100,
      null,
      null,
      "Job site delivery",
    ],
    [
      new Date(today.getTime() - 2 * 24 * 60 * 60 * 1000),
      "A-002",
      null,
      "Refuel",
      "Mike Johnson",
      "Costco",
      38150,
      22.3,
      71.2,
      "Regular unleaded",
    ],
    [
      new Date(today.getTime() - 3 * 24 * 60 * 60 * 1000),
      "A-003",
      null,
      "Check-out",
      "Tom Brown",
      "Main Yard",
      1250,
      null,
      null,
      "Excavation job",
    ],
    [
      new Date(today.getTime() - 3 * 24 * 60 * 60 * 1000),
      "A-003",
      null,
      "Refuel",
      "Tom Brown",
      "On-site",
      1258,
      15.0,
      58.5,
      "Diesel",
    ],
    [
      new Date(today.getTime() - 3 * 24 * 60 * 60 * 1000),
      "A-003",
      null,
      "Check-in",
      "Tom Brown",
      "Main Yard",
      1262,
      null,
      null,
      "Job complete",
    ],
    [
      new Date(today.getTime() - 4 * 24 * 60 * 60 * 1000),
      "A-006",
      null,
      "Check-out",
      "Sarah Wilson",
      "Main Yard",
      125,
      null,
      null,
      "Cleaning job",
    ],
    [
      new Date(today.getTime() - 5 * 24 * 60 * 60 * 1000),
      "A-009",
      null,
      "Check-out",
      "Tom Brown",
      "Main Yard",
      890,
      null,
      null,
      "Landscaping project",
    ],
    [
      new Date(today.getTime() - 5 * 24 * 60 * 60 * 1000),
      "A-009",
      null,
      "Refuel",
      "Tom Brown",
      "Main Yard",
      898,
      12.0,
      48.0,
      "Diesel",
    ],
    [
      new Date(today.getTime() - 6 * 24 * 60 * 60 * 1000),
      "A-001",
      null,
      "Refuel",
      "John Smith",
      "BP Station",
      45150,
      17.2,
      58.25,
      "Regular unleaded",
    ],
    [
      new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000),
      "A-007",
      null,
      "Check-out",
      "Mike Johnson",
      "Main Yard",
      12500,
      null,
      null,
      "Parts pickup",
    ],
    [
      new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000),
      "A-007",
      null,
      "Check-in",
      "Mike Johnson",
      "Main Yard",
      12545,
      null,
      null,
      "Returned",
    ],
    [
      new Date(today.getTime() - 10 * 24 * 60 * 60 * 1000),
      "A-002",
      null,
      "Refuel",
      "Mike Johnson",
      "Shell Station",
      38050,
      20.1,
      68.34,
      "Regular unleaded",
    ],
  ];

  activity.getRange(2, 1, activityData.length, 10).setValues(activityData);

  // Add Asset Name lookup formula
  for (let i = 2; i <= activityData.length + 1; i++) {
    activity
      .getRange(i, 3)
      .setFormula(
        "=IFERROR(VLOOKUP(B" + i + ",'Assets Master'!A:B,2,FALSE),\"\")",
      );
  }

  // ============================================
  // MAINTENANCE TRACKER DATA
  // ============================================
  const maint = ss.getSheetByName("Maintenance Tracker");

  const maintData = [
    [
      "M-001",
      "A-001",
      null,
      "Oil Change",
      new Date(2024, 10, 1),
      new Date(2024, 10, 1),
      null,
      85,
      "Quick Lube",
      "Synthetic oil",
    ],
    [
      "M-002",
      "A-002",
      null,
      "Tire Rotation",
      new Date(2024, 9, 15),
      new Date(2024, 9, 15),
      null,
      45,
      "Discount Tire",
      "All 4 tires",
    ],
    [
      "M-003",
      "A-004",
      null,
      "Repair",
      new Date(2024, 11, 10),
      null,
      null,
      0,
      "In-house",
      "Blade replacement",
    ],
    [
      "M-004",
      "A-003",
      null,
      "Annual Inspection",
      new Date(2024, 11, 1),
      new Date(2024, 11, 1),
      null,
      350,
      "CAT Dealer",
      "Passed",
    ],
    [
      "M-005",
      "A-005",
      null,
      "Inspection",
      new Date(2024, 6, 15),
      new Date(2024, 6, 15),
      null,
      75,
      "DOT Station",
      "Annual trailer",
    ],
    [
      "M-006",
      "A-001",
      null,
      "Oil Change",
      new Date(today.getTime() + 5 * 24 * 60 * 60 * 1000),
      null,
      null,
      0,
      "Quick Lube",
      "Scheduled",
    ],
    [
      "M-007",
      "A-002",
      null,
      "Oil Change",
      new Date(today.getTime() + 3 * 24 * 60 * 60 * 1000),
      null,
      null,
      0,
      "Quick Lube",
      "Due soon",
    ],
    [
      "M-008",
      "A-009",
      null,
      "Filter Replacement",
      new Date(today.getTime() - 5 * 24 * 60 * 60 * 1000),
      null,
      null,
      0,
      "In-house",
      "OVERDUE!",
    ],
    [
      "M-009",
      "A-006",
      null,
      "Inspection",
      new Date(2024, 11, 10),
      new Date(2024, 11, 10),
      null,
      50,
      "In-house",
      "Pre-season check",
    ],
    [
      "M-010",
      "A-007",
      null,
      "Oil Change",
      new Date(2024, 10, 20),
      new Date(2024, 10, 20),
      null,
      95,
      "Jiffy Lube",
      "Synthetic blend",
    ],
  ];

  maint.getRange(2, 1, maintData.length, 10).setValues(maintData);

  // Add formulas for Asset Name and Status
  for (let i = 2; i <= maintData.length + 1; i++) {
    maint
      .getRange(i, 3)
      .setFormula(
        "=IFERROR(VLOOKUP(B" + i + ",'Assets Master'!A:B,2,FALSE),\"\")",
      );
    maint
      .getRange(i, 7)
      .setFormula(
        "=IF(ISBLANK(E" +
          i +
          '),"",IF(NOT(ISBLANK(F' +
          i +
          ')),"Completed",IF(E' +
          i +
          '<TODAY(),"Overdue","Scheduled")))',
      );
  }

  // ============================================
  // COST TRACKING DATA
  // ============================================
  const cost = ss.getSheetByName("Cost Tracking");

  // Clear any existing data
  if (cost.getLastRow() > 1) {
    cost.getRange(2, 1, cost.getLastRow(), 7).clearContent();
  }

  // Add each asset with formulas
  for (let i = 0; i < assetData.length; i++) {
    const row = i + 2;
    const assetId = assetData[i][0];

    cost.getRange(row, 1).setValue(assetId);
    cost
      .getRange(row, 2)
      .setFormula(
        "=IFERROR(VLOOKUP(A" + row + ",'Assets Master'!A:B,2,FALSE),\"\")",
      );
    cost
      .getRange(row, 3)
      .setFormula(
        "=SUMIFS('Activity Log'!I:I,'Activity Log'!B:B,A" +
          row +
          ",'Activity Log'!A:A,\">=\"&TODAY()-30)",
      );
    cost
      .getRange(row, 4)
      .setFormula(
        "=SUMIFS('Maintenance Tracker'!H:H,'Maintenance Tracker'!B:B,A" +
          row +
          ",'Maintenance Tracker'!F:F,\">=\"&DATE(YEAR(TODAY()),1,1))",
      );
    cost
      .getRange(row, 5)
      .setFormula(
        "=COUNTIFS('Activity Log'!B:B,A" +
          row +
          ",'Activity Log'!D:D,\"Check-out\")",
      );
    cost
      .getRange(row, 6)
      .setFormula("=IFERROR(IF(E" + row + ">0,C" + row + "/E" + row + ",0),0)");
    cost
      .getRange(row, 7)
      .setFormula(
        "=IFERROR(IF(ISBLANK(VLOOKUP(A" +
          row +
          ",'Assets Master'!A:G,7,FALSE)),\"\",TODAY()-VLOOKUP(A" +
          row +
          ",'Assets Master'!A:G,7,FALSE)),\"\")",
      );
  }

  // ============================================
  // CONFIG DATA - Only set if empty (don't overwrite user's settings)
  // ============================================
  const config = ss.getSheetByName("Config");
  if (!config.getRange("B2").getValue()) {
    config.getRange("B2").setValue("test");
  }
  if (!config.getRange("B3").getValue()) {
    config.getRange("B3").setValue("jacob@truenorthstrategyops.com");
  }

  // Force recalculation
  SpreadsheetApp.flush();

  ui.alert(
    "✅ Test Data Added!\n\n" +
      "• 10 assets\n" +
      "• 15 activity entries\n" +
      "• 10 maintenance records\n" +
      "• Cost tracking populated\n\n" +
      "Check your Dashboards now!",
  );

  Logger.log("Test data added successfully");
}

/**
 * Clears all test data while preserving structure
 */
function clearTestData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    "Clear All Data?",
    "This will remove all data from Assets, Activity, Maintenance, and Cost Tracking sheets. Continue?",
    ui.ButtonSet.YES_NO,
  );

  if (response !== ui.Button.YES) return;

  // Clear data sheets (preserve headers)
  const sheets = [
    "Assets Master",
    "Activity Log",
    "Maintenance Tracker",
    "Cost Tracking",
  ];
  sheets.forEach((name) => {
    const sheet = ss.getSheetByName(name);
    if (sheet && sheet.getLastRow() > 1) {
      sheet
        .getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn())
        .clearContent();
    }
  });

  // Reset Config
  const config = ss.getSheetByName("Config");
  if (config) {
    config.getRange("B2").setValue("");
    config.getRange("B3").setValue("");
  }

  SpreadsheetApp.flush();

  ui.alert("✅ All data cleared!\n\nHeaders and structure preserved.");
  Logger.log("Test data cleared");
}

// ============================================
// EMAIL AUTOMATION FUNCTIONS
// ============================================

/**
 * Sends daily digest email with asset statistics
 */
function sendDailyDigest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Get email from config
  const config = ss.getSheetByName("Config");
  const email = config?.getRange("B3").getValue();
  const businessName = config?.getRange("B2").getValue() || "G-Stack Asset Command";

  if (!email) {
    ui.alert(
      "No email configured!\n\nPlease add your email in Config sheet cell B3.",
    );
    return;
  }

  // Get asset counts
  const assets = ss.getSheetByName("Assets Master");
  const data = assets.getDataRange().getValues();

  let total = 0,
    available = 0,
    inUse = 0,
    maintenance = 0,
    retired = 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      total++;
      const status = data[i][3];
      if (status === "Available") available++;
      else if (status === "In Use") inUse++;
      else if (status === "Maintenance") maintenance++;
      else if (status === "Retired") retired++;
    }
  }

  // Check for maintenance due
  const today = new Date();
  const threshold = config.getRange("B4").getValue() || 7;
  let maintenanceDue = 0;

  for (let i = 1; i < data.length; i++) {
    const nextService = data[i][7];
    if (nextService instanceof Date) {
      const daysUntil = Math.floor(
        (nextService - today) / (1000 * 60 * 60 * 24),
      );
      if (daysUntil <= threshold && daysUntil >= 0) maintenanceDue++;
    }
  }

  // Build email
  const subject = `${businessName} Daily Digest - ${today.toLocaleDateString()}`;
  const body = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <h2 style="color: #0F172A; border-bottom: 2px solid #00B4D8; padding-bottom: 10px;">
        ${businessName} - Daily Asset Summary
      </h2>

      <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
        <tr>
          <td style="padding: 12px; background: #3B82F6; color: white; font-weight: bold; text-align: center;">Total Assets</td>
          <td style="padding: 12px; background: #10B981; color: white; font-weight: bold; text-align: center;">Available</td>
          <td style="padding: 12px; background: #F59E0B; color: white; font-weight: bold; text-align: center;">In Use</td>
          <td style="padding: 12px; background: #EF4444; color: white; font-weight: bold; text-align: center;">Maintenance</td>
        </tr>
        <tr>
          <td style="padding: 12px; text-align: center; font-size: 24px; font-weight: bold;">${total}</td>
          <td style="padding: 12px; text-align: center; font-size: 24px; font-weight: bold;">${available}</td>
          <td style="padding: 12px; text-align: center; font-size: 24px; font-weight: bold;">${inUse}</td>
          <td style="padding: 12px; text-align: center; font-size: 24px; font-weight: bold;">${maintenance}</td>
        </tr>
      </table>

      ${
        maintenanceDue > 0
          ? `
        <div style="background: #FEF3C7; border-left: 4px solid #F59E0B; padding: 15px; margin: 20px 0;">
          <strong>⚠️ Maintenance Alert:</strong> ${maintenanceDue} asset(s) have maintenance due within ${threshold} days.
        </div>
      `
          : ""
      }

      <p style="margin-top: 20px;">
        <a href="${ss.getUrl()}" style="background: #00B4D8; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">
          Open G-Stack Asset Command
        </a>
      </p>

      <p style="color: #666; font-size: 12px; margin-top: 30px;">
        Sent by G-Stack Asset Command Free Starter Kit | True North Data Strategies
      </p>
    </div>
  `;

  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: body,
  });

  ui.alert(`✅ Daily digest sent to ${email}`);
  Logger.log("Daily digest sent to " + email);
}

/**
 * Checks for upcoming maintenance and sends alert email
 */
function checkMaintenanceDue() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Get config
  const config = ss.getSheetByName("Config");
  const email = config?.getRange("B3").getValue();
  const businessName = config?.getRange("B2").getValue() || "G-Stack Asset Command";
  const threshold = config?.getRange("B4").getValue() || 7;

  if (!email) {
    ui.alert(
      "No email configured!\n\nPlease add your email in Config sheet cell B3.",
    );
    return;
  }

  // Get assets and check maintenance dates
  const assets = ss.getSheetByName("Assets Master");
  const data = assets.getDataRange().getValues();
  const today = new Date();
  const alerts = [];

  for (let i = 1; i < data.length; i++) {
    const assetId = data[i][0];
    const assetName = data[i][1];
    const nextService = data[i][7];

    if (assetId && nextService instanceof Date) {
      const daysUntil = Math.floor(
        (nextService - today) / (1000 * 60 * 60 * 24),
      );

      if (daysUntil <= threshold) {
        alerts.push({
          id: assetId,
          name: assetName,
          dueDate: nextService.toLocaleDateString(),
          daysUntil: daysUntil,
          urgent: daysUntil <= 3,
        });
      }
    }
  }

  // Also check Maintenance Tracker for overdue
  const maint = ss.getSheetByName("Maintenance Tracker");
  const maintData = maint.getDataRange().getValues();

  for (let i = 1; i < maintData.length; i++) {
    const status = maintData[i][6];
    if (status === "Overdue") {
      const exists = alerts.find((a) => a.id === maintData[i][1]);
      if (!exists) {
        alerts.push({
          id: maintData[i][1],
          name: maintData[i][2] || "Unknown",
          dueDate:
            maintData[i][4] instanceof Date
              ? maintData[i][4].toLocaleDateString()
              : "Unknown",
          daysUntil: -1,
          urgent: true,
        });
      }
    }
  }

  if (alerts.length === 0) {
    ui.alert(`✅ No maintenance due within ${threshold} days!`);
    return;
  }

  // Sort by urgency
  alerts.sort((a, b) => a.daysUntil - b.daysUntil);

  // Build email
  const subject = `⚠️ ${businessName}: ${alerts.length} Maintenance Alert(s)`;

  let tableRows = alerts
    .map(
      (a) => `
    <tr style="${a.urgent ? "background: #FEE2E2;" : ""}">
      <td style="padding: 8px; border: 1px solid #ddd;">${a.id}</td>
      <td style="padding: 8px; border: 1px solid #ddd;">${a.name}</td>
      <td style="padding: 8px; border: 1px solid #ddd;">${a.dueDate}</td>
      <td style="padding: 8px; border: 1px solid #ddd; text-align: center; font-weight: bold; ${a.daysUntil < 0 ? "color: #EF4444;" : ""}">${a.daysUntil < 0 ? "OVERDUE" : a.daysUntil + " days"}</td>
    </tr>
  `,
    )
    .join("");

  const body = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <h2 style="color: #EF4444;">⚠️ Maintenance Alerts</h2>
      <p>${alerts.length} asset(s) need attention:</p>

      <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
        <tr style="background: #1E3A5F; color: white;">
          <th style="padding: 10px; text-align: left;">Asset ID</th>
          <th style="padding: 10px; text-align: left;">Name</th>
          <th style="padding: 10px; text-align: left;">Due Date</th>
          <th style="padding: 10px; text-align: center;">Time Left</th>
        </tr>
        ${tableRows}
      </table>

      <p>
        <a href="${ss.getUrl()}" style="background: #00B4D8; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">
          Open G-Stack Asset Command
        </a>
      </p>

      <p style="color: #666; font-size: 12px; margin-top: 30px;">
        Sent by G-Stack Asset Command Free Starter Kit | True North Data Strategies
      </p>
    </div>
  `;

  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: body,
  });

  ui.alert(
    `✅ Maintenance alerts sent to ${email}\n\n${alerts.length} asset(s) need attention.`,
  );
  Logger.log("Maintenance alerts sent: " + alerts.length);
}

// ============================================
// HTML DASHBOARD FUNCTIONS
// ============================================

/**
 * Show the interactive HTML dashboard
 */
function showDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const template = HtmlService.createTemplateFromFile("Dashboard");
  template.spreadsheetId = ss ? ss.getId() : "";

  const html = template
    .evaluate()
    .setWidth(1200)
    .setHeight(800)
    .setTitle("G-Stack Asset Command Dashboard");
  SpreadsheetApp.getUi().showModalDialog(html, "🚀 G-Stack Asset Command Dashboard");
}

/**
 * Get all dashboard data for the HTML dashboard
 */
function getDashboardData(spreadsheetId) {
  const ss = getAssetCommandSpreadsheet_(spreadsheetId);
  const config = ss.getSheetByName("Config");
  let businessName = "G-Stack Asset Command";

  if (config) {
    const configData = config.getDataRange().getValues();
    for (let i = 0; i < configData.length; i++) {
      if (configData[i][0] === "Business Name") {
        businessName = configData[i][1] || businessName;
        break;
      }
    }
  }

  return {
    companyName: businessName,
    lastUpdated: new Date().toLocaleString(),
    assets: getAssetStats(ss),
    activity: getActivityStats(ss),
    maintenance: getMaintenanceStats(ss),
    costs: getCostStats(ss),
    charts: getAssetChartData(ss),
  };
}

/**
 * Get asset statistics
 */
function getAssetStats(ss) {
  const sheet = ss.getSheetByName("Assets Master");
  const stats = {
    total: 0,
    available: 0,
    inUse: 0,
    maintenance: 0,
    retired: 0,
    byType: {},
    totalValue: 0,
  };

  if (!sheet) return stats;

  try {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue; // Skip empty rows

      stats.total++;
      const type = (data[i][2] || "Other").toString();
      const status = (data[i][3] || "").toString();
      const cost = parseFloat(data[i][10]) || 0;

      stats.totalValue += cost;
      stats.byType[type] = (stats.byType[type] || 0) + 1;

      switch (status) {
        case "Available":
          stats.available++;
          break;
        case "In Use":
          stats.inUse++;
          break;
        case "Maintenance":
          stats.maintenance++;
          break;
        case "Retired":
          stats.retired++;
          break;
      }
    }
  } catch (e) {
    Logger.log("Error reading assets: " + e.message);
  }

  return stats;
}

/**
 * Get activity statistics
 */
function getActivityStats(ss) {
  const sheet = ss.getSheetByName("Activity Log");
  const stats = {
    total: 0,
    today: 0,
    thisWeek: 0,
    checkOuts: 0,
    checkIns: 0,
    refuels: 0,
    maintenanceActions: 0,
    recentActivity: [],
  };

  if (!sheet) return stats;

  try {
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    const todayStart = new Date(
      now.getFullYear(),
      now.getMonth(),
      now.getDate(),
    );
    const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;

      stats.total++;
      const timestamp = new Date(data[i][0]);
      const action = (data[i][3] || "").toString();

      if (timestamp >= todayStart) stats.today++;
      if (timestamp >= weekAgo) stats.thisWeek++;

      switch (action) {
        case "Check-out":
          stats.checkOuts++;
          break;
        case "Check-in":
          stats.checkIns++;
          break;
        case "Refuel":
          stats.refuels++;
          break;
        case "Maintenance":
          stats.maintenanceActions++;
          break;
      }

      // Collect recent activity (last 5)
      if (stats.recentActivity.length < 5 && i <= 5) {
        stats.recentActivity.push({
          timestamp: timestamp.toLocaleString(),
          assetId: data[i][1],
          assetName: data[i][2],
          action: action,
          employee: data[i][4],
        });
      }
    }
  } catch (e) {
    Logger.log("Error reading activity: " + e.message);
  }

  return stats;
}

/**
 * Get maintenance statistics
 */
function getMaintenanceStats(ss) {
  const sheet = ss.getSheetByName("Maintenance Tracker");
  const stats = {
    total: 0,
    pending: 0,
    completed: 0,
    overdue: 0,
    upcoming: [],
    overdueList: [],
  };

  if (!sheet) return stats;

  try {
    const data = sheet.getDataRange().getValues();
    const now = new Date();

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;

      stats.total++;
      const scheduledDate = data[i][4] ? new Date(data[i][4]) : null;
      const completedDate = data[i][5];
      const status = (data[i][6] || "").toString();

      if (completedDate || status === "Completed") {
        stats.completed++;
      } else {
        stats.pending++;
        if (scheduledDate) {
          const daysUntil = Math.ceil(
            (scheduledDate - now) / (1000 * 60 * 60 * 24),
          );

          if (daysUntil < 0) {
            stats.overdue++;
            if (stats.overdueList.length < 5) {
              stats.overdueList.push({
                assetId: data[i][1],
                assetName: data[i][2],
                serviceType: data[i][3],
                dueDate: scheduledDate.toLocaleDateString(),
                daysOverdue: Math.abs(daysUntil),
              });
            }
          } else if (daysUntil <= 14 && stats.upcoming.length < 5) {
            stats.upcoming.push({
              assetId: data[i][1],
              assetName: data[i][2],
              serviceType: data[i][3],
              dueDate: scheduledDate.toLocaleDateString(),
              daysUntil: daysUntil,
            });
          }
        }
      }
    }

    stats.upcoming.sort((a, b) => a.daysUntil - b.daysUntil);
    stats.overdueList.sort((a, b) => b.daysOverdue - a.daysOverdue);
  } catch (e) {
    Logger.log("Error reading maintenance: " + e.message);
  }

  return stats;
}

/**
 * Get cost statistics
 */
function getCostStats(ss) {
  const stats = {
    totalCosts: 0,
    fuelCosts: 0,
    maintenanceCosts: 0,
    otherCosts: 0,
    ytdCosts: 0,
  };

  try {
    const currentYear = new Date().getFullYear();

    // Primary source: Activity Log (fuel) + Maintenance Tracker (maintenance).
    // This matches the template structure users actually work in.
    const activity = ss.getSheetByName("Activity Log");
    if (activity) {
      const activityData = activity.getDataRange().getValues();
      for (let i = 1; i < activityData.length; i++) {
        if (!activityData[i][0]) continue;
        if (activityData[i][3] !== "Refuel") continue;

        const date = new Date(activityData[i][0]);
        const fuelCost = parseFloat(activityData[i][8]) || 0;

        if (fuelCost <= 0) continue;
        stats.fuelCosts += fuelCost;
        if (date.getFullYear() === currentYear) {
          stats.ytdCosts += fuelCost;
        }
      }
    }

    const maintenance = ss.getSheetByName("Maintenance Tracker");
    if (maintenance) {
      const maintData = maintenance.getDataRange().getValues();
      for (let i = 1; i < maintData.length; i++) {
        if (!maintData[i][0]) continue;

        const scheduledDate = maintData[i][4];
        const completedDate = maintData[i][5];
        const date =
          completedDate instanceof Date ? completedDate : scheduledDate;
        const totalCost = parseFloat(maintData[i][9]) || 0;

        if (!(date instanceof Date) || totalCost <= 0) continue;
        stats.maintenanceCosts += totalCost;
        if (date.getFullYear() === currentYear) {
          stats.ytdCosts += totalCost;
        }
      }
    }

    // Backward-compatible support for older ledger-style Cost Tracking sheets.
    const costSheet = ss.getSheetByName("Cost Tracking");
    if (costSheet) {
      const costData = costSheet.getDataRange().getValues();
      const headers = (costData[0] || []).map((h) =>
        (h || "").toString().toLowerCase(),
      );
      const hasLedgerColumns =
        headers.includes("category") && headers.includes("amount");

      if (hasLedgerColumns) {
        const dateIdx = headers.indexOf("date");
        const categoryIdx = headers.indexOf("category");
        const amountIdx = headers.indexOf("amount");

        for (let i = 1; i < costData.length; i++) {
          const amount = parseFloat(costData[i][amountIdx]) || 0;
          if (amount <= 0) continue;

          const date = new Date(costData[i][dateIdx]);
          const category = (costData[i][categoryIdx] || "")
            .toString()
            .toLowerCase();

          if (category.includes("fuel")) {
            stats.fuelCosts += amount;
          } else if (
            category.includes("maintenance") ||
            category.includes("repair")
          ) {
            stats.maintenanceCosts += amount;
          } else {
            stats.otherCosts += amount;
          }

          if (date.getFullYear() === currentYear) {
            stats.ytdCosts += amount;
          }
        }
      }
    }

    stats.totalCosts =
      stats.fuelCosts + stats.maintenanceCosts + stats.otherCosts;
  } catch (e) {
    Logger.log("Error reading costs: " + e.message);
  }

  return stats;
}

/**
 * Get chart data for the dashboard
 */
function getAssetChartData(ss) {
  const assets = getAssetStats(ss);

  const types = Object.entries(assets.byType).sort((a, b) => b[1] - a[1]);

  return {
    assetTypes: types.map((t) => t[0]),
    assetTypeCounts: types.map((t) => t[1]),
    statusLabels: ["Available", "In Use", "Maintenance", "Retired"],
    statusCounts: [
      assets.available,
      assets.inUse,
      assets.maintenance,
      assets.retired,
    ],
  };
}

/**
 * Execute a dashboard action
 */
function executeDashboardAction(action, spreadsheetId) {
  try {
    const ss = getAssetCommandSpreadsheet_(spreadsheetId);
    const sid = ss.getId();

    switch (action) {
      case "refresh":
        SpreadsheetApp.flush();
        return { success: true, message: "Dashboards refreshed" };
      case "checkMaintenance":
        checkMaintenanceDueFromDashboard(sid);
        return {
          success: true,
          message: "Maintenance check complete - email sent if alerts found",
        };
      case "checkCompliance":
        checkDriverComplianceFromDashboard(sid);
        return {
          success: true,
          message: "Compliance check complete - email sent if alerts found",
        };
      case "checkFuel":
        checkFuelAnomalyFromDashboard(sid);
        return {
          success: true,
          message: "Fuel check complete - email sent if anomalies found",
        };
      case "sendDigest":
        sendDailyDigestFromDashboard(sid);
        return { success: true, message: "Daily digest email sent" };
      default:
        return { success: false, message: "Unknown action" };
    }
  } catch (error) {
    Logger.log("Dashboard action error: " + error.message);
    return { success: false, message: error.message };
  }
}

/**
 * Dashboard version of checkDriverCompliance (no UI alerts)
 */
function checkDriverComplianceFromDashboard(spreadsheetId) {
  const ss = getAssetCommandSpreadsheet_(spreadsheetId);
  const config = ss.getSheetByName("Config");
  const email = config?.getRange("B3").getValue();
  const businessName = config?.getRange("B2").getValue() || "G-Stack Asset Command";
  const threshold = 30;

  if (!email) throw new Error("No email configured in Config sheet");

  const drivers = ss.getSheetByName("Drivers");
  if (!drivers) throw new Error("Drivers sheet not found");

  const data = drivers.getDataRange().getValues();
  const today = new Date();
  const alerts = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] || data[i][2] === "Terminated") continue;
    const driverId = data[i][0];
    const driverName = data[i][1];

    // Check License, CDL, Medical
    [
      [7, "License"],
      [10, "CDL"],
      [12, "Medical Card"],
    ].forEach(([idx, type]) => {
      const expiry = data[i][idx];
      if (expiry instanceof Date) {
        const daysUntil = Math.floor((expiry - today) / (1000 * 60 * 60 * 24));
        if (daysUntil <= threshold && daysUntil > -30) {
          alerts.push({
            driver: driverName,
            driverId,
            type,
            expiry: expiry.toLocaleDateString(),
            daysUntil,
            urgent: daysUntil <= 7,
          });
        }
      }
    });

    // Clearinghouse
    const clearinghouse = data[i][16];
    if (clearinghouse === "Violation" || clearinghouse === "Pending Query") {
      alerts.push({
        driver: driverName,
        driverId,
        type: "Clearinghouse",
        expiry: clearinghouse,
        daysUntil: -999,
        urgent: true,
      });
    }
  }

  if (alerts.length === 0) return;

  alerts.sort((a, b) => a.daysUntil - b.daysUntil);
  const subject =
    businessName + ": " + alerts.length + " Driver Compliance Alert(s)";
  let tableRows = "";
  for (const a of alerts) {
    const urgentStyle = a.urgent ? "background: #FEE2E2;" : "";
    const statusText =
      a.daysUntil === -999
        ? "ACTION REQUIRED"
        : a.daysUntil < 0
          ? "EXPIRED"
          : a.daysUntil + " days";
    tableRows +=
      '<tr style="' +
      urgentStyle +
      '"><td style="padding: 8px; border: 1px solid #ddd;">' +
      a.driverId +
      '</td><td style="padding: 8px; border: 1px solid #ddd;">' +
      a.driver +
      '</td><td style="padding: 8px; border: 1px solid #ddd;">' +
      a.type +
      '</td><td style="padding: 8px; border: 1px solid #ddd;">' +
      a.expiry +
      '</td><td style="padding: 8px; border: 1px solid #ddd; font-weight: bold;">' +
      statusText +
      "</td></tr>";
  }
  const body =
    '<div style="font-family: Arial, sans-serif;"><h2 style="color: #EF4444;">Driver Compliance Alerts</h2><table style="width: 100%; border-collapse: collapse;"><tr style="background: #1E3A5F; color: white;"><th style="padding: 10px;">ID</th><th style="padding: 10px;">Name</th><th style="padding: 10px;">Credential</th><th style="padding: 10px;">Expires</th><th style="padding: 10px;">Status</th></tr>' +
    tableRows +
    "</table></div>";
  MailApp.sendEmail({ to: email, subject: subject, htmlBody: body });
}

/**
 * Dashboard version of checkFuelAnomaly (no UI alerts)
 */
function checkFuelAnomalyFromDashboard(spreadsheetId) {
  const ss = getAssetCommandSpreadsheet_(spreadsheetId);
  const config = ss.getSheetByName("Config");
  const email = config?.getRange("B3").getValue();
  const businessName = config?.getRange("B2").getValue() || "G-Stack Asset Command";
  const anomalyThreshold = config?.getRange("B5").getValue() || 20;

  if (!email) throw new Error("No email configured");

  const activity = ss.getSheetByName("Activity Log");
  if (!activity) return;

  const data = activity.getDataRange().getValues();
  const fuelByAsset = {};
  const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);

  for (let i = 1; i < data.length; i++) {
    if (data[i][3] !== "Refuel") continue;
    const date = new Date(data[i][0]);
    if (date < thirtyDaysAgo) continue;
    const assetId = data[i][1];
    const gallons = parseFloat(data[i][7]) || 0;
    if (!fuelByAsset[assetId]) fuelByAsset[assetId] = { entries: [], total: 0 };
    fuelByAsset[assetId].entries.push(gallons);
    fuelByAsset[assetId].total += gallons;
  }

  const anomalies = [];
  for (const assetId in fuelByAsset) {
    const asset = fuelByAsset[assetId];
    if (asset.entries.length < 3) continue;
    const avg = asset.total / asset.entries.length;
    for (const entry of asset.entries) {
      const variance = Math.abs((entry - avg) / avg) * 100;
      if (variance > anomalyThreshold) {
        anomalies.push({
          assetId,
          amount: entry,
          average: avg.toFixed(1),
          variance: variance.toFixed(1),
        });
        break;
      }
    }
  }

  if (anomalies.length === 0) return;

  const subject = businessName + ": Unusual Fuel Usage Detected";
  let tableRows = "";
  for (const a of anomalies) {
    tableRows +=
      '<tr><td style="padding: 8px; border: 1px solid #ddd;">' +
      a.assetId +
      '</td><td style="padding: 8px; border: 1px solid #ddd;">' +
      a.amount +
      ' gal</td><td style="padding: 8px; border: 1px solid #ddd;">' +
      a.average +
      ' gal</td><td style="padding: 8px; border: 1px solid #ddd; color: #EF4444; font-weight: bold;">' +
      a.variance +
      "%</td></tr>";
  }
  const body =
    '<div style="font-family: Arial, sans-serif;"><h2 style="color: #F59E0B;">Fuel Anomaly Alert</h2><table style="width: 100%; border-collapse: collapse;"><tr style="background: #1E3A5F; color: white;"><th style="padding: 10px;">Asset</th><th style="padding: 10px;">Fill</th><th style="padding: 10px;">Avg</th><th style="padding: 10px;">Variance</th></tr>' +
    tableRows +
    "</table></div>";
  MailApp.sendEmail({ to: email, subject: subject, htmlBody: body });
}

// ============================================
// DRIVER COMPLIANCE FUNCTIONS
// ============================================

/**
 * Checks for expiring driver credentials and sends alerts
 */
function checkDriverCompliance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const config = ss.getSheetByName("Config");
  const email = config?.getRange("B3").getValue();
  const businessName = config?.getRange("B2").getValue() || "G-Stack Asset Command";
  const threshold = 30; // 30 days for license alerts

  if (!email) {
    ui.alert("No email configured! Add email in Config sheet cell B3.");
    return;
  }

  const drivers = ss.getSheetByName("Drivers");
  if (!drivers) {
    ui.alert("Drivers sheet not found. Run Build Complete Template first.");
    return;
  }

  const data = drivers.getDataRange().getValues();
  const today = new Date();
  const alerts = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] || data[i][2] === "Terminated") continue;

    const driverId = data[i][0];
    const driverName = data[i][1];

    // Check License Expiry (column H, index 7)
    const licenseExpiry = data[i][7];
    if (licenseExpiry instanceof Date) {
      const daysUntil = Math.floor(
        (licenseExpiry - today) / (1000 * 60 * 60 * 24),
      );
      if (daysUntil <= threshold && daysUntil > -30) {
        alerts.push({
          driver: driverName,
          driverId: driverId,
          type: "License",
          expiry: licenseExpiry.toLocaleDateString(),
          daysUntil: daysUntil,
          urgent: daysUntil <= 7,
        });
      }
    }

    // Check CDL Expiry (column K, index 10)
    const cdlExpiry = data[i][10];
    if (cdlExpiry instanceof Date) {
      const daysUntil = Math.floor((cdlExpiry - today) / (1000 * 60 * 60 * 24));
      if (daysUntil <= threshold && daysUntil > -30) {
        alerts.push({
          driver: driverName,
          driverId: driverId,
          type: "CDL",
          expiry: cdlExpiry.toLocaleDateString(),
          daysUntil: daysUntil,
          urgent: daysUntil <= 7,
        });
      }
    }

    // Check Medical Card Expiry (column M, index 12)
    const medicalExpiry = data[i][12];
    if (medicalExpiry instanceof Date) {
      const daysUntil = Math.floor(
        (medicalExpiry - today) / (1000 * 60 * 60 * 24),
      );
      if (daysUntil <= threshold && daysUntil > -30) {
        alerts.push({
          driver: driverName,
          driverId: driverId,
          type: "Medical Card",
          expiry: medicalExpiry.toLocaleDateString(),
          daysUntil: daysUntil,
          urgent: daysUntil <= 14,
        });
      }
    }

    // Check Clearinghouse Status (column Q, index 16)
    const clearinghouse = data[i][16];
    if (clearinghouse === "Violation" || clearinghouse === "Pending Query") {
      alerts.push({
        driver: driverName,
        driverId: driverId,
        type: "Clearinghouse",
        expiry: clearinghouse,
        daysUntil: -999,
        urgent: true,
      });
    }
  }

  if (alerts.length === 0) {
    ui.alert("All driver credentials are current!");
    return;
  }

  alerts.sort((a, b) => a.daysUntil - b.daysUntil);

  const subject =
    businessName + ": " + alerts.length + " Driver Compliance Alert(s)";

  let tableRows = "";
  for (const a of alerts) {
    const urgentStyle = a.urgent ? "background: #FEE2E2;" : "";
    const statusText =
      a.daysUntil === -999
        ? "ACTION REQUIRED"
        : a.daysUntil < 0
          ? "EXPIRED"
          : a.daysUntil + " days";
    const statusStyle = a.daysUntil < 0 ? "color: #EF4444;" : "";
    tableRows +=
      '<tr style="' +
      urgentStyle +
      '">' +
      '<td style="padding: 8px; border: 1px solid #ddd;">' +
      a.driverId +
      "</td>" +
      '<td style="padding: 8px; border: 1px solid #ddd;">' +
      a.driver +
      "</td>" +
      '<td style="padding: 8px; border: 1px solid #ddd;">' +
      a.type +
      "</td>" +
      '<td style="padding: 8px; border: 1px solid #ddd;">' +
      a.expiry +
      "</td>" +
      '<td style="padding: 8px; border: 1px solid #ddd; text-align: center; font-weight: bold; ' +
      statusStyle +
      '">' +
      statusText +
      "</td></tr>";
  }

  const body =
    '<div style="font-family: Arial, sans-serif; max-width: 700px;">' +
    '<h2 style="color: #EF4444;">Driver Compliance Alerts</h2>' +
    "<p>" +
    alerts.length +
    " credential(s) need attention:</p>" +
    '<table style="width: 100%; border-collapse: collapse;">' +
    '<tr style="background: #1E3A5F; color: white;">' +
    '<th style="padding: 10px;">ID</th><th style="padding: 10px;">Name</th>' +
    '<th style="padding: 10px;">Credential</th><th style="padding: 10px;">Expires</th>' +
    '<th style="padding: 10px;">Status</th></tr>' +
    tableRows +
    "</table>" +
    '<p style="margin-top: 20px;"><a href="' +
    ss.getUrl() +
    '" style="background: #00B4D8; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">Open G-Stack Asset Command</a></p></div>';

  MailApp.sendEmail({ to: email, subject: subject, htmlBody: body });
  ui.alert("Driver compliance alerts sent to " + email);
  Logger.log("Driver compliance alerts sent: " + alerts.length);
}

/**
 * Checks for unusual fuel usage and sends alert
 */
function checkFuelAnomaly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const config = ss.getSheetByName("Config");
  const email = config?.getRange("B3").getValue();
  const businessName = config?.getRange("B2").getValue() || "G-Stack Asset Command";
  const anomalyThreshold = config?.getRange("B5").getValue() || 20;

  if (!email) {
    ui.alert("No email configured!");
    return;
  }

  const activity = ss.getSheetByName("Activity Log");
  if (!activity) return;

  const data = activity.getDataRange().getValues();
  const fuelByAsset = {};
  const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);

  for (let i = 1; i < data.length; i++) {
    if (data[i][3] !== "Refuel") continue;
    const date = new Date(data[i][0]);
    if (date < thirtyDaysAgo) continue;

    const assetId = data[i][1];
    const gallons = parseFloat(data[i][7]) || 0;

    if (!fuelByAsset[assetId]) fuelByAsset[assetId] = { entries: [], total: 0 };
    fuelByAsset[assetId].entries.push(gallons);
    fuelByAsset[assetId].total += gallons;
  }

  const anomalies = [];
  for (const assetId in fuelByAsset) {
    const asset = fuelByAsset[assetId];
    if (asset.entries.length < 3) continue;
    const avg = asset.total / asset.entries.length;
    for (const entry of asset.entries) {
      const variance = Math.abs((entry - avg) / avg) * 100;
      if (variance > anomalyThreshold) {
        anomalies.push({
          assetId: assetId,
          amount: entry,
          average: avg.toFixed(1),
          variance: variance.toFixed(1),
        });
        break;
      }
    }
  }

  if (anomalies.length === 0) {
    ui.alert("No fuel anomalies detected!");
    return;
  }

  const subject = businessName + ": Unusual Fuel Usage Detected";
  let tableRows = "";
  for (const a of anomalies) {
    tableRows +=
      '<tr><td style="padding: 8px; border: 1px solid #ddd;">' +
      a.assetId +
      "</td>" +
      '<td style="padding: 8px; border: 1px solid #ddd;">' +
      a.amount +
      " gal</td>" +
      '<td style="padding: 8px; border: 1px solid #ddd;">' +
      a.average +
      " gal</td>" +
      '<td style="padding: 8px; border: 1px solid #ddd; color: #EF4444; font-weight: bold;">' +
      a.variance +
      "%</td></tr>";
  }

  const body =
    '<div style="font-family: Arial, sans-serif;">' +
    '<h2 style="color: #F59E0B;">Fuel Anomaly Alert</h2>' +
    "<p>Unusual fuel consumption (>" +
    anomalyThreshold +
    "% variance):</p>" +
    '<table style="width: 100%; border-collapse: collapse;">' +
    '<tr style="background: #1E3A5F; color: white;">' +
    '<th style="padding: 10px;">Asset</th><th style="padding: 10px;">Fill</th>' +
    '<th style="padding: 10px;">Avg</th><th style="padding: 10px;">Variance</th></tr>' +
    tableRows +
    "</table>" +
    '<p style="margin-top: 20px;"><a href="' +
    ss.getUrl() +
    '" style="background: #00B4D8; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">Open G-Stack Asset Command</a></p></div>';

  MailApp.sendEmail({ to: email, subject: subject, htmlBody: body });
  ui.alert("Fuel anomaly alerts sent to " + email);
}

/**
 * Add sample driver data for testing
 */
function addDriverTestData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const drivers = ss.getSheetByName("Drivers");

  if (!drivers) {
    SpreadsheetApp.getUi().alert("Drivers sheet not found.");
    return;
  }

  const today = new Date();

  // Test data with all 30 columns including HOS/ELD/DVIR
  const driverData = [
    [
      "DRV-001",
      "John Smith",
      "Active",
      "(555) 123-4567",
      "jsmith@example.com",
      "DL123456",
      "TX",
      new Date(today.getTime() + 180 * 24 * 60 * 60 * 1000),
      "Class A",
      "CDL789012",
      new Date(today.getTime() + 365 * 24 * 60 * 60 * 1000),
      "H,N,T",
      new Date(today.getTime() + 90 * 24 * 60 * 60 * 1000),
      "Dr. Johnson",
      new Date(today.getTime() - 90 * 24 * 60 * 60 * 1000),
      new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000),
      "Clear",
      new Date(today.getTime() - 60 * 24 * 60 * 60 * 1000),
      new Date(today.getTime() - 15 * 24 * 60 * 60 * 1000),
      "Negative",
      new Date(today.getTime() - 45 * 24 * 60 * 60 * 1000),
      "KeepTruckin",
      "KT-001234",
      "Yes",
      new Date(today.getTime() - 1 * 24 * 60 * 60 * 1000),
      new Date(today.getTime() - 1 * 24 * 60 * 60 * 1000),
      new Date(today.getTime() + 180 * 24 * 60 * 60 * 1000),
      new Date(2022, 5, 15),
      "",
      "Senior driver - fully compliant",
    ],

    [
      "DRV-002",
      "Mike Johnson",
      "Active",
      "(555) 234-5678",
      "mjohnson@example.com",
      "DL234567",
      "TX",
      new Date(today.getTime() + 15 * 24 * 60 * 60 * 1000),
      "Class A",
      "CDL890123",
      new Date(today.getTime() + 45 * 24 * 60 * 60 * 1000),
      "H,N",
      new Date(today.getTime() + 20 * 24 * 60 * 60 * 1000),
      "Dr. Williams",
      new Date(today.getTime() - 180 * 24 * 60 * 60 * 1000),
      new Date(today.getTime() - 60 * 24 * 60 * 60 * 1000),
      "Clear",
      new Date(today.getTime() - 90 * 24 * 60 * 60 * 1000),
      new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000),
      "Negative",
      new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000),
      "Samsara",
      "SM-005678",
      "Yes",
      new Date(today.getTime() - 2 * 24 * 60 * 60 * 1000),
      new Date(today.getTime() - 2 * 24 * 60 * 60 * 1000),
      new Date(today.getTime() + 90 * 24 * 60 * 60 * 1000),
      new Date(2023, 2, 10),
      "",
      "License expiring soon",
    ],

    [
      "DRV-003",
      "Sarah Wilson",
      "Active",
      "(555) 345-6789",
      "swilson@example.com",
      "DL345678",
      "TX",
      new Date(today.getTime() + 365 * 24 * 60 * 60 * 1000),
      "Class B",
      "CDL901234",
      new Date(today.getTime() + 365 * 24 * 60 * 60 * 1000),
      "P",
      new Date(today.getTime() - 10 * 24 * 60 * 60 * 1000),
      "Dr. Brown",
      new Date(today.getTime() - 200 * 24 * 60 * 60 * 1000),
      new Date(today.getTime() - 45 * 24 * 60 * 60 * 1000),
      "Clear",
      new Date(today.getTime() - 120 * 24 * 60 * 60 * 1000),
      new Date(today.getTime() - 60 * 24 * 60 * 60 * 1000),
      "Negative",
      new Date(today.getTime() - 60 * 24 * 60 * 60 * 1000),
      "Geotab",
      "GT-009012",
      "Yes",
      new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000),
      new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000),
      new Date(today.getTime() + 30 * 24 * 60 * 60 * 1000),
      new Date(2023, 8, 1),
      "",
      "Medical EXPIRED - needs renewal",
    ],

    [
      "DRV-004",
      "Tom Brown",
      "On Leave",
      "(555) 456-7890",
      "tbrown@example.com",
      "DL456789",
      "TX",
      new Date(today.getTime() + 200 * 24 * 60 * 60 * 1000),
      "Class A",
      "CDL012345",
      new Date(today.getTime() + 300 * 24 * 60 * 60 * 1000),
      "H,N,T,X",
      new Date(today.getTime() + 150 * 24 * 60 * 60 * 1000),
      "Dr. Davis",
      new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000),
      new Date(today.getTime() - 90 * 24 * 60 * 60 * 1000),
      "Pending Query",
      new Date(today.getTime() - 180 * 24 * 60 * 60 * 1000),
      new Date(today.getTime() - 20 * 24 * 60 * 60 * 1000),
      "Negative",
      new Date(today.getTime() - 90 * 24 * 60 * 60 * 1000),
      "Omnitracs",
      "OM-003456",
      "Pending",
      new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000),
      new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000),
      new Date(today.getTime() - 15 * 24 * 60 * 60 * 1000),
      new Date(2021, 1, 20),
      "",
      "Clearinghouse issue + Annual inspection overdue",
    ],
  ];

  drivers.getRange(2, 1, driverData.length, 30).setValues(driverData);
  SpreadsheetApp.getUi().alert(
    "Driver test data added! 4 drivers with HOS/ELD/DVIR compliance scenarios.",
  );
}

/**
 * Clear driver test data
 */
function clearDriverTestData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const drivers = ss.getSheetByName("Drivers");
  if (drivers && drivers.getLastRow() > 1) {
    drivers
      .getRange(2, 1, drivers.getLastRow(), drivers.getLastColumn())
      .clearContent();
    SpreadsheetApp.getUi().alert("Driver data cleared!");
  }
}

/**
 * Dashboard-safe digest sender (no Spreadsheet UI calls).
 */
function sendDailyDigestFromDashboard(spreadsheetId) {
  const ss = getAssetCommandSpreadsheet_(spreadsheetId);
  const config = ss.getSheetByName("Config");
  const email = config?.getRange("B3").getValue();
  const businessName = config?.getRange("B2").getValue() || "G-Stack Asset Command";

  if (!email) throw new Error("No email configured in Config sheet");

  const assets = ss.getSheetByName("Assets Master");
  if (!assets) throw new Error("Assets Master sheet not found");

  const data = assets.getDataRange().getValues();
  let total = 0,
    available = 0,
    inUse = 0,
    maintenance = 0;

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    total++;
    const status = data[i][3];
    if (status === "Available") available++;
    else if (status === "In Use") inUse++;
    else if (status === "Maintenance") maintenance++;
  }

  const today = new Date();
  const threshold = config.getRange("B4").getValue() || 7;
  let maintenanceDue = 0;

  for (let i = 1; i < data.length; i++) {
    const nextService = data[i][7];
    if (nextService instanceof Date) {
      const daysUntil = Math.floor(
        (nextService - today) / (1000 * 60 * 60 * 24),
      );
      if (daysUntil <= threshold && daysUntil >= 0) maintenanceDue++;
    }
  }

  const subject =
    businessName + " Daily Digest - " + today.toLocaleDateString();
  const body =
    '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">' +
    '<h2 style="color: #0F172A; border-bottom: 2px solid #00B4D8; padding-bottom: 10px;">' +
    businessName +
    " - Daily Asset Summary</h2>" +
    '<table style="width: 100%; border-collapse: collapse; margin: 20px 0;">' +
    '<tr><td style="padding: 12px; background: #3B82F6; color: white; font-weight: bold; text-align: center;">Total Assets</td>' +
    '<td style="padding: 12px; background: #10B981; color: white; font-weight: bold; text-align: center;">Available</td>' +
    '<td style="padding: 12px; background: #F59E0B; color: white; font-weight: bold; text-align: center;">In Use</td>' +
    '<td style="padding: 12px; background: #EF4444; color: white; font-weight: bold; text-align: center;">Maintenance</td></tr>' +
    '<tr><td style="padding: 12px; text-align: center; font-size: 24px; font-weight: bold;">' +
    total +
    "</td>" +
    '<td style="padding: 12px; text-align: center; font-size: 24px; font-weight: bold;">' +
    available +
    "</td>" +
    '<td style="padding: 12px; text-align: center; font-size: 24px; font-weight: bold;">' +
    inUse +
    "</td>" +
    '<td style="padding: 12px; text-align: center; font-size: 24px; font-weight: bold;">' +
    maintenance +
    "</td></tr></table>" +
    (maintenanceDue > 0
      ? '<div style="background: #FEF3C7; border-left: 4px solid #F59E0B; padding: 15px; margin: 20px 0;"><strong>⚠️ Maintenance Alert:</strong> ' +
        maintenanceDue +
        " asset(s) have maintenance due within " +
        threshold +
        " days.</div>"
      : "") +
    '<p><a href="' +
    ss.getUrl() +
    '" style="background: #00B4D8; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">Open G-Stack Asset Command</a></p>' +
    "</div>";

  MailApp.sendEmail({ to: email, subject: subject, htmlBody: body });
}

/**
 * Dashboard-safe maintenance check (no Spreadsheet UI calls).
 */
function checkMaintenanceDueFromDashboard(spreadsheetId) {
  const ss = getAssetCommandSpreadsheet_(spreadsheetId);
  const config = ss.getSheetByName("Config");
  const email = config?.getRange("B3").getValue();
  const businessName = config?.getRange("B2").getValue() || "G-Stack Asset Command";
  const threshold = config?.getRange("B4").getValue() || 7;

  if (!email) throw new Error("No email configured in Config sheet");

  const assets = ss.getSheetByName("Assets Master");
  if (!assets) throw new Error("Assets Master sheet not found");

  const data = assets.getDataRange().getValues();
  const today = new Date();
  const alerts = [];

  for (let i = 1; i < data.length; i++) {
    const assetId = data[i][0];
    const assetName = data[i][1];
    const nextService = data[i][7];

    if (assetId && nextService instanceof Date) {
      const daysUntil = Math.floor(
        (nextService - today) / (1000 * 60 * 60 * 24),
      );
      if (daysUntil <= threshold) {
        alerts.push({
          id: assetId,
          name: assetName,
          dueDate: nextService.toLocaleDateString(),
          daysUntil: daysUntil,
          urgent: daysUntil <= 3,
        });
      }
    }
  }

  const maint = ss.getSheetByName("Maintenance Tracker");
  if (maint) {
    const maintData = maint.getDataRange().getValues();
    for (let i = 1; i < maintData.length; i++) {
      if (maintData[i][6] === "Overdue") {
        const exists = alerts.find((a) => a.id === maintData[i][1]);
        if (!exists) {
          alerts.push({
            id: maintData[i][1],
            name: maintData[i][2] || "Unknown",
            dueDate:
              maintData[i][4] instanceof Date
                ? maintData[i][4].toLocaleDateString()
                : "Unknown",
            daysUntil: -1,
            urgent: true,
          });
        }
      }
    }
  }

  if (alerts.length === 0) return;

  alerts.sort((a, b) => a.daysUntil - b.daysUntil);
  const subject =
    "⚠️ " + businessName + ": " + alerts.length + " Maintenance Alert(s)";
  const rows = alerts
    .map(
      (a) =>
        '<tr style="' +
        (a.urgent ? "background: #FEE2E2;" : "") +
        '">' +
        '<td style="padding: 8px; border: 1px solid #ddd;">' +
        a.id +
        "</td>" +
        '<td style="padding: 8px; border: 1px solid #ddd;">' +
        a.name +
        "</td>" +
        '<td style="padding: 8px; border: 1px solid #ddd;">' +
        a.dueDate +
        "</td>" +
        '<td style="padding: 8px; border: 1px solid #ddd; text-align: center; font-weight: bold; ' +
        (a.daysUntil < 0 ? "color:#EF4444;" : "") +
        '">' +
        (a.daysUntil < 0 ? "OVERDUE" : a.daysUntil + " days") +
        "</td></tr>",
    )
    .join("");

  const body =
    '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">' +
    '<h2 style="color: #EF4444;">⚠️ Maintenance Alerts</h2>' +
    "<p>" +
    alerts.length +
    " asset(s) need attention:</p>" +
    '<table style="width: 100%; border-collapse: collapse; margin: 20px 0;">' +
    '<tr style="background: #1E3A5F; color: white;"><th style="padding:10px; text-align:left;">Asset ID</th><th style="padding:10px; text-align:left;">Name</th><th style="padding:10px; text-align:left;">Due Date</th><th style="padding:10px; text-align:center;">Time Left</th></tr>' +
    rows +
    "</table></div>";

  MailApp.sendEmail({ to: email, subject: subject, htmlBody: body });
}

// ============================================
// DASHBOARD DATA FUNCTIONS
// ============================================

/**
 * Get compliance dashboard data for the HTML dashboard
 */
function getComplianceDashboardData(spreadsheetId) {
  const ss = getAssetCommandSpreadsheet_(spreadsheetId);
  const drivers = ss.getSheetByName("Drivers");

  const result = {
    activeDrivers: 0,
    expiredCount: 0,
    expiringSoonCount: 0,
    clearinghouseIssues: 0,
    expiredList: [],
    expiringList: [],
    drivers: [],
  };

  if (!drivers) return result;

  try {
    const data = drivers.getDataRange().getValues();
    const today = new Date();
    const threshold = 30; // days

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;

      const driverId = data[i][0];
      const driverName = data[i][1];
      const status = data[i][2];

      // Skip terminated drivers
      if (status === "Terminated") continue;
      if (status === "Active") result.activeDrivers++;

      const licenseExpiry = data[i][7];
      const cdlExpiry = data[i][10];
      const medicalExpiry = data[i][12];
      const clearinghouse = data[i][16];

      // Build driver record for table
      const driverRecord = {
        name: driverName,
        licenseStatus: "na",
        licenseDays: null,
        cdlStatus: "na",
        cdlDays: null,
        medicalStatus: "na",
        medicalDays: null,
        clearinghouse: clearinghouse || "N/A",
      };

      // Check License
      if (licenseExpiry instanceof Date) {
        const daysUntil = Math.floor(
          (licenseExpiry - today) / (1000 * 60 * 60 * 24),
        );
        if (daysUntil < 0) {
          driverRecord.licenseStatus = "expired";
          result.expiredCount++;
          result.expiredList.push({
            driverName: driverName,
            credentialType: "License",
            expiredDate: licenseExpiry.toLocaleDateString(),
          });
        } else if (daysUntil <= threshold) {
          driverRecord.licenseStatus = "expiring";
          driverRecord.licenseDays = daysUntil;
          result.expiringSoonCount++;
          result.expiringList.push({
            driverName: driverName,
            credentialType: "License",
            expiryDate: licenseExpiry.toLocaleDateString(),
            daysUntil: daysUntil,
          });
        } else {
          driverRecord.licenseStatus = "ok";
        }
      }

      // Check CDL
      if (cdlExpiry instanceof Date) {
        const daysUntil = Math.floor(
          (cdlExpiry - today) / (1000 * 60 * 60 * 24),
        );
        if (daysUntil < 0) {
          driverRecord.cdlStatus = "expired";
          result.expiredCount++;
          result.expiredList.push({
            driverName: driverName,
            credentialType: "CDL",
            expiredDate: cdlExpiry.toLocaleDateString(),
          });
        } else if (daysUntil <= threshold) {
          driverRecord.cdlStatus = "expiring";
          driverRecord.cdlDays = daysUntil;
          result.expiringSoonCount++;
          result.expiringList.push({
            driverName: driverName,
            credentialType: "CDL",
            expiryDate: cdlExpiry.toLocaleDateString(),
            daysUntil: daysUntil,
          });
        } else {
          driverRecord.cdlStatus = "ok";
        }
      }

      // Check Medical Card
      if (medicalExpiry instanceof Date) {
        const daysUntil = Math.floor(
          (medicalExpiry - today) / (1000 * 60 * 60 * 24),
        );
        if (daysUntil < 0) {
          driverRecord.medicalStatus = "expired";
          result.expiredCount++;
          result.expiredList.push({
            driverName: driverName,
            credentialType: "Medical Card",
            expiredDate: medicalExpiry.toLocaleDateString(),
          });
        } else if (daysUntil <= threshold) {
          driverRecord.medicalStatus = "expiring";
          driverRecord.medicalDays = daysUntil;
          result.expiringSoonCount++;
          result.expiringList.push({
            driverName: driverName,
            credentialType: "Medical Card",
            expiryDate: medicalExpiry.toLocaleDateString(),
            daysUntil: daysUntil,
          });
        } else {
          driverRecord.medicalStatus = "ok";
        }
      }

      // Check Clearinghouse
      if (clearinghouse === "Violation" || clearinghouse === "Pending Query") {
        result.clearinghouseIssues++;
      }

      result.drivers.push(driverRecord);
    }

    // Sort lists
    result.expiredList.sort((a, b) => a.driverName.localeCompare(b.driverName));
    result.expiringList.sort((a, b) => a.daysUntil - b.daysUntil);
  } catch (e) {
    Logger.log("Error getting compliance data: " + e.message);
  }

  return result;
}

/**
 * Get fuel dashboard data for the HTML dashboard
 */
function getFuelDashboardData(spreadsheetId) {
  const ss = getAssetCommandSpreadsheet_(spreadsheetId);
  const activity = ss.getSheetByName("Activity Log");
  const assets = ss.getSheetByName("Assets Master");
  const config = ss.getSheetByName("Config");

  const result = {
    totalGallons: 0,
    totalCost: 0,
    avgCostPerGallon: 0,
    anomalyCount: 0,
    anomalies: [],
    recentRefuels: [],
    topAssetsByFuel: [],
    topAssetsByCost: [],
  };

  if (!activity) return result;

  const anomalyThreshold = config?.getRange("B5").getValue() || 20;

  try {
    const activityData = activity.getDataRange().getValues();
    const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
    const fuelByAsset = {};
    const recentRefuels = [];

    // Get asset names lookup
    const assetNames = {};
    if (assets) {
      const assetData = assets.getDataRange().getValues();
      for (let i = 1; i < assetData.length; i++) {
        if (assetData[i][0]) {
          assetNames[assetData[i][0]] = assetData[i][1];
        }
      }
    }

    // Process activity log
    for (let i = 1; i < activityData.length; i++) {
      if (activityData[i][3] !== "Refuel") continue;

      const date = new Date(activityData[i][0]);
      const assetId = activityData[i][1];
      const assetName = activityData[i][2] || assetNames[assetId] || assetId;
      const employee = activityData[i][4];
      const gallons = parseFloat(activityData[i][7]) || 0;
      const cost = parseFloat(activityData[i][8]) || 0;

      // Only count last 30 days
      if (date >= thirtyDaysAgo) {
        result.totalGallons += gallons;
        result.totalCost += cost;

        // Track by asset
        if (!fuelByAsset[assetId]) {
          fuelByAsset[assetId] = {
            name: assetName,
            gallons: 0,
            cost: 0,
            entries: [],
          };
        }
        fuelByAsset[assetId].gallons += gallons;
        fuelByAsset[assetId].cost += cost;
        fuelByAsset[assetId].entries.push(gallons);

        // Recent refuels (top 10)
        if (recentRefuels.length < 10) {
          recentRefuels.push({
            assetId: assetId,
            assetName: assetName,
            date: date.toLocaleDateString(),
            employee: employee,
            gallons: gallons.toFixed(1),
            cost: cost.toFixed(2),
          });
        }
      }
    }

    // Calculate average cost per gallon
    if (result.totalGallons > 0) {
      result.avgCostPerGallon = result.totalCost / result.totalGallons;
    }

    // Detect anomalies
    for (const assetId in fuelByAsset) {
      const asset = fuelByAsset[assetId];
      if (asset.entries.length < 3) continue;

      const avg = asset.gallons / asset.entries.length;
      for (const entry of asset.entries) {
        const variance = Math.abs((entry - avg) / avg) * 100;
        if (variance > anomalyThreshold) {
          result.anomalyCount++;
          result.anomalies.push({
            assetId: assetId,
            assetName: asset.name,
            amount: entry.toFixed(1),
            average: avg.toFixed(1),
            variance: variance.toFixed(1),
          });
          break; // One anomaly per asset
        }
      }
    }

    // Top assets by fuel
    const sortedByFuel = Object.values(fuelByAsset)
      .sort((a, b) => b.gallons - a.gallons)
      .slice(0, 5);
    result.topAssetsByFuel = sortedByFuel.map((a) => ({
      name: a.name,
      gallons: Math.round(a.gallons),
    }));

    // Top assets by cost
    const sortedByCost = Object.values(fuelByAsset)
      .sort((a, b) => b.cost - a.cost)
      .slice(0, 5);
    result.topAssetsByCost = sortedByCost.map((a) => ({
      name: a.name,
      cost: Math.round(a.cost),
    }));

    result.recentRefuels = recentRefuels;
  } catch (e) {
    Logger.log("Error getting fuel data: " + e.message);
  }

  return result;
}

// ============================================
// SIDEBAR FUNCTIONS
// ============================================

/**
 * Opens the data entry sidebar
 */
function openDataEntrySidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("Data Entry")
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Opens the user manual dialog
 */
function showUserManual() {
  const html = HtmlService.createHtmlOutputFromFile("Help")
    .setTitle("G-Stack Asset Command User Manual")
    .setWidth(950)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, "G-Stack Asset Command User Manual");
}

/**
 * Add a single asset from sidebar form
 */
function addAsset(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let assets = ss.getSheetByName("Assets Master");

    if (!assets) {
      return {
        success: false,
        message: "Assets Master sheet not found. Run Build Template first.",
      };
    }

    // Check for duplicate Asset ID
    const existingData = assets.getDataRange().getValues();
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][0] === data.assetId) {
        return {
          success: false,
          message: "Asset ID already exists: " + data.assetId,
        };
      }
    }

    // Calculate next service date
    let nextService = "";
    if (data.lastService && data.serviceInterval) {
      const lastDate = new Date(data.lastService);
      nextService = new Date(
        lastDate.getTime() + data.serviceInterval * 24 * 60 * 60 * 1000,
      );
    }

    // Prepare row data
    const row = [
      data.assetId,
      data.assetName,
      data.assetType,
      data.status,
      data.location || "",
      data.assignedTo || "",
      data.lastService ? new Date(data.lastService) : "",
      data.serviceInterval || "",
      nextService,
      "", // Current Odometer
      "", // Last Fuel Date
      "", // Fuel Tank Size
      data.acquisitionDate ? new Date(data.acquisitionDate) : "",
      data.acquisitionCost || "",
      "", // Depreciation Rate
      "", // Current Value
      data.notes || "",
      new Date(), // Date Added
    ];

    assets.appendRow(row);
    return { success: true, message: "Asset added: " + data.assetId };
  } catch (e) {
    Logger.log("Error adding asset: " + e.message);
    return { success: false, message: "Error: " + e.message };
  }
}

/**
 * Add a single driver from sidebar form
 */
function addDriver(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let drivers = ss.getSheetByName("Drivers");

    if (!drivers) {
      return {
        success: false,
        message: "Drivers sheet not found. Run Build Template first.",
      };
    }

    // Check for duplicate Driver ID
    const existingData = drivers.getDataRange().getValues();
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][0] === data.driverId) {
        return {
          success: false,
          message: "Driver ID already exists: " + data.driverId,
        };
      }
    }

    // Prepare row data (30 columns for Drivers sheet)
    const row = [
      data.driverId,
      data.driverName,
      data.status,
      data.phone || "",
      data.email || "",
      data.hireDate ? new Date(data.hireDate) : "",
      data.licenseNumber || "",
      data.licenseExpiry ? new Date(data.licenseExpiry) : "",
      data.licenseState || "",
      data.cdlNumber || "",
      data.cdlExpiry ? new Date(data.cdlExpiry) : "",
      data.licenseType || "",
      data.medicalExpiry ? new Date(data.medicalExpiry) : "",
      "", // MVR Review Date
      "", // Safety Training Date
      "", // Background Check Date
      "", // Clearinghouse Status
      "", // HOS Status
      "", // HOS Last Check
      "", // ELD Compliance
      "", // DVIR Status
      "", // DVIR Last Check
      "", // Accident History
      "", // Safety Score
      "", // Endorsements
      "", // Restrictions
      "", // Emergency Contact Name
      "", // Emergency Contact Phone
      data.notes || "",
      new Date(), // Date Added
    ];

    drivers.appendRow(row);
    return { success: true, message: "Driver added: " + data.driverName };
  } catch (e) {
    Logger.log("Error adding driver: " + e.message);
    return { success: false, message: "Error: " + e.message };
  }
}

/**
 * Add a maintenance record from sidebar form
 */
function addMaintenance(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let maint = ss.getSheetByName("Maintenance Tracker");

    if (!maint) {
      return {
        success: false,
        message:
          "Maintenance Tracker sheet not found. Run Build Template first.",
      };
    }

    // Get asset name
    const assets = ss.getSheetByName("Assets Master");
    let assetName = "";
    if (assets) {
      const assetData = assets.getDataRange().getValues();
      for (let i = 1; i < assetData.length; i++) {
        if (assetData[i][0] === data.assetId) {
          assetName = assetData[i][1];
          break;
        }
      }
    }

    // Calculate total cost
    const partsCost = parseFloat(data.partsCost) || 0;
    const laborCost = parseFloat(data.laborCost) || 0;
    const totalCost = partsCost + laborCost;

    // Determine status
    let status = "Scheduled";
    if (data.completedDate) {
      status = "Completed";
    }

    // Prepare row data
    const row = [
      data.maintId,
      data.assetId,
      assetName,
      data.serviceType,
      data.scheduledDate ? new Date(data.scheduledDate) : "",
      data.completedDate ? new Date(data.completedDate) : "",
      status,
      partsCost,
      laborCost,
      totalCost,
      data.vendor || "",
      data.invoiceNumber || "",
      data.notes || "",
      new Date(), // Date Created
    ];

    maint.appendRow(row);
    return {
      success: true,
      message: "Maintenance record added: " + data.maintId,
    };
  } catch (e) {
    Logger.log("Error adding maintenance: " + e.message);
    return { success: false, message: "Error: " + e.message };
  }
}

/**
 * Add an activity/cost record from sidebar form
 */
function addActivityRecord(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let activity = ss.getSheetByName("Activity Log");

    if (!activity) {
      return {
        success: false,
        message: "Activity Log sheet not found. Run Build Template first.",
      };
    }

    // Get asset name
    const assets = ss.getSheetByName("Assets Master");
    let assetName = "";
    if (assets) {
      const assetData = assets.getDataRange().getValues();
      for (let i = 1; i < assetData.length; i++) {
        if (assetData[i][0] === data.assetId) {
          assetName = assetData[i][1];
          break;
        }
      }
    }

    // Prepare row data
    const row = [
      new Date(), // Timestamp
      data.assetId,
      assetName,
      data.actionType,
      data.employee || "",
      data.location || "",
      data.odometer || "",
      data.actionType === "Refuel" ? data.fuelGallons || "" : "",
      data.actionType === "Refuel" ? data.fuelCost || "" : "",
      data.notes || "",
    ];

    activity.appendRow(row);

    // Update asset status if check-in/check-out
    if (
      assets &&
      (data.actionType === "Check-out" || data.actionType === "Check-in")
    ) {
      const assetData = assets.getDataRange().getValues();
      for (let i = 1; i < assetData.length; i++) {
        if (assetData[i][0] === data.assetId) {
          const newStatus =
            data.actionType === "Check-out" ? "In Use" : "Available";
          assets.getRange(i + 1, 4).setValue(newStatus);
          if (data.actionType === "Check-out") {
            assets.getRange(i + 1, 6).setValue(data.employee || "");
            assets.getRange(i + 1, 5).setValue(data.location || "");
          } else {
            assets.getRange(i + 1, 6).setValue("");
          }
          break;
        }
      }
    }

    return { success: true, message: "Activity record added" };
  } catch (e) {
    Logger.log("Error adding activity: " + e.message);
    return { success: false, message: "Error: " + e.message };
  }
}

/**
 * Import CSV data
 */
function importCsvData(type, csvData) {
  try {
    const rows = parseCsv(csvData);
    if (rows.length < 2) {
      return {
        success: false,
        message: "CSV file is empty or has no data rows",
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let count = 0;

    switch (type) {
      case "assets":
        count = importAssetsFromCsv(ss, rows);
        break;
      case "drivers":
        count = importDriversFromCsv(ss, rows);
        break;
      case "maintenance":
        count = importMaintenanceFromCsv(ss, rows);
        break;
      case "costs":
        count = importCostsFromCsv(ss, rows);
        break;
      default:
        return { success: false, message: "Unknown import type: " + type };
    }

    return {
      success: true,
      message: "Imported " + count + " records successfully!",
    };
  } catch (e) {
    Logger.log("Error importing CSV: " + e.message);
    return { success: false, message: "Error: " + e.message };
  }
}

/**
 * Parse CSV string into array of arrays
 */
function parseCsv(csvString) {
  const rows = [];
  const lines = csvString.split(/\r?\n/);

  for (const line of lines) {
    if (line.trim() === "") continue;

    const row = [];
    let inQuotes = false;
    let currentValue = "";

    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      const nextChar = line[i + 1];

      if (char === '"') {
        if (inQuotes && nextChar === '"') {
          currentValue += '"';
          i++;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (char === "," && !inQuotes) {
        row.push(currentValue.trim());
        currentValue = "";
      } else {
        currentValue += char;
      }
    }
    row.push(currentValue.trim());
    rows.push(row);
  }

  return rows;
}

/**
 * Import assets from CSV rows
 */
function importAssetsFromCsv(ss, rows) {
  const assets = ss.getSheetByName("Assets Master");
  if (!assets) throw new Error("Assets Master sheet not found");

  let count = 0;
  // Skip header row
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;

    const row = [
      r[0] || "", // Asset ID
      r[1] || "", // Asset Name
      r[2] || "Vehicle", // Type
      r[3] || "Available", // Status
      r[4] || "", // Location
      r[5] || "", // Assigned To
      r[6] ? new Date(r[6]) : "", // Last Service
      r[7] || "", // Service Interval
      "", // Next Service (calculated)
      r[8] || "", // Odometer
      "", // Last Fuel Date
      r[9] || "", // Fuel Tank Size
      r[10] ? new Date(r[10]) : "", // Acquisition Date
      r[11] || "", // Acquisition Cost
      "",
      "", // Depreciation
      r[12] || "", // Notes
      new Date(),
    ];
    assets.appendRow(row);
    count++;
  }
  return count;
}

/**
 * Import drivers from CSV rows
 */
function importDriversFromCsv(ss, rows) {
  const drivers = ss.getSheetByName("Drivers");
  if (!drivers) throw new Error("Drivers sheet not found");

  let count = 0;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;

    const row = [
      r[0] || "", // Driver ID
      r[1] || "", // Name
      r[2] || "Active", // Status
      r[3] || "", // Phone
      r[4] || "", // Email
      r[5] ? new Date(r[5]) : "", // Hire Date
      r[6] || "", // License Number
      r[7] ? new Date(r[7]) : "", // License Expiry
      r[8] || "", // License State
      r[9] || "", // CDL Number
      r[10] ? new Date(r[10]) : "", // CDL Expiry
      r[11] || "", // License Type
      r[12] ? new Date(r[12]) : "", // Medical Expiry
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      r[13] || "", // Notes
      new Date(),
    ];
    drivers.appendRow(row);
    count++;
  }
  return count;
}

/**
 * Import maintenance from CSV rows
 */
function importMaintenanceFromCsv(ss, rows) {
  const maint = ss.getSheetByName("Maintenance Tracker");
  if (!maint) throw new Error("Maintenance Tracker sheet not found");

  let count = 0;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;

    const partsCost = parseFloat(r[5]) || 0;
    const laborCost = parseFloat(r[6]) || 0;

    const row = [
      r[0] || "", // Maintenance ID
      r[1] || "", // Asset ID
      r[2] || "", // Asset Name
      r[3] || "", // Service Type
      r[4] ? new Date(r[4]) : "", // Scheduled Date
      r[7] ? new Date(r[7]) : "", // Completed Date
      r[8] || "Scheduled", // Status
      partsCost,
      laborCost,
      partsCost + laborCost,
      r[9] || "", // Vendor
      r[10] || "", // Invoice Number
      r[11] || "", // Notes
      new Date(),
    ];
    maint.appendRow(row);
    count++;
  }
  return count;
}

/**
 * Import cost/activity records from CSV rows
 */
function importCostsFromCsv(ss, rows) {
  const activity = ss.getSheetByName("Activity Log");
  if (!activity) throw new Error("Activity Log sheet not found");

  let count = 0;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;

    const row = [
      r[0] ? new Date(r[0]) : new Date(), // Timestamp
      r[1] || "", // Asset ID
      r[2] || "", // Asset Name
      r[3] || "", // Action Type
      r[4] || "", // Employee
      r[5] || "", // Location
      r[6] || "", // Odometer
      r[7] || "", // Fuel Gallons
      r[8] || "", // Fuel Cost
      r[9] || "", // Notes
    ];
    activity.appendRow(row);
    count++;
  }
  return count;
}

/**
 * Get CSV template for download
 */
function getCsvTemplate(type) {
  let template = "";

  switch (type) {
    case "assets":
      template =
        "Asset ID,Asset Name,Type,Status,Location,Assigned To,Last Service Date,Service Interval Days,Odometer,Fuel Tank Size,Acquisition Date,Acquisition Cost,Notes\n";
      template +=
        "A-001,Red F-150,Vehicle,Available,Main Yard,,2024-01-15,90,45000,26,2022-06-01,35000,Company truck\n";
      template +=
        "A-002,Trailer #1,Trailer,Available,Main Yard,,,,,,,20000,Flatbed trailer\n";
      break;

    case "drivers":
      template =
        "Driver ID,Driver Name,Status,Phone,Email,Hire Date,License Number,License Expiry,License State,CDL Number,CDL Expiry,License Type,Medical Expiry,Notes\n";
      template +=
        "DRV-001,John Smith,Active,(555) 123-4567,john@example.com,2023-01-15,DL123456,2026-06-15,TX,CDL789012,2025-06-15,Class A,2025-03-01,Experienced driver\n";
      break;

    case "maintenance":
      template =
        "Maintenance ID,Asset ID,Asset Name,Service Type,Scheduled Date,Parts Cost,Labor Cost,Completed Date,Status,Vendor,Invoice Number,Notes\n";
      template +=
        "M-001,A-001,Red F-150,Oil Change,2024-02-15,45.00,30.00,2024-02-15,Completed,Quick Lube,INV-12345,Regular maintenance\n";
      template +=
        "M-002,A-001,Red F-150,Tire Rotation,2024-03-01,0,25.00,,Scheduled,,,Due next month\n";
      break;

    case "costs":
      template =
        "Date,Asset ID,Asset Name,Action Type,Employee,Location,Odometer,Fuel Gallons,Fuel Cost,Notes\n";
      template +=
        "2024-01-15,A-001,Red F-150,Check-out,John Smith,Main Yard,45000,,,Picked up for delivery\n";
      template +=
        "2024-01-15,A-001,Red F-150,Refuel,John Smith,Shell Station,45150,18.5,62.50,Regular fill-up\n";
      template +=
        "2024-01-15,A-001,Red F-150,Check-in,John Smith,Main Yard,45200,,,Returned after delivery\n";
      break;

    default:
      return { success: false, message: "Unknown template type" };
  }

  return { success: true, template: template };
}
