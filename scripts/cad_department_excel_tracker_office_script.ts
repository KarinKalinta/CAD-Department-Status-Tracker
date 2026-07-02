function main(workbook: ExcelScript.Workbook) {
  const planName = "Plan of Action";
  const examplesName = "CAD Examples";
  const listsName = "Lists hidden";

  const headers = [
    "WE",
    "Request Type",
    "Main Category",
    "Sub Task",
    "Deliverable Type",
    "Estimated Time",
    "Priority",
    "Status",
    "Owner",
    "Start Date",
    "Due Date",
    "Notes",
    "Link / File"
  ];

  const requestTypes = ["Initial Request", "New Request", "Revision", "Follow-up"];

  const mainCategories = [
    "C0553-F",
    "C0731 Mireyas",
    "C0634-X",
    "C0634",
    "Organic",
    "Men's",
    "Cluster",
    "RhinoArtisan",
    "PD CAD Investigation",
    "Cosmic Collection",
    "Small Scale Design",
    "Custom Order",
    "Internal CAD",
    "Other"
  ];

  const subTasks = [
    "Sketch",
    "Existing sketch",
    "3D set",
    "Photo",
    "CAD",
    "CAD Render Files",
    "CAD PD Files",
    "Presentation",
    "PD Update Progress",
    "Render",
    "Sample",
    "R&D",
    "Finding Data",
    "Update Production to CCU and OPC",
    "Test with real job"
  ];

  const deliverableTypes = [
    "Sketch",
    "Existing Sketch",
    "3D Asset",
    "Photo",
    "CAD File",
    "Email",
    "Render",
    "Sample",
    "Document",
    "Presentation",
    "STL / Print File",
    "Other"
  ];

  const estimatedTimes = [
    "15 minutes",
    "20 minutes",
    "30 to 45 minutes",
    "1 hour",
    "2 to 3 hours",
    "Half day",
    "1 day",
    "2 days",
    "This week"
  ];

  const priorities = ["High", "Medium", "Low"];
  const statuses = ["In Progress", "Waiting", "Done", "Blocked", "Cancelled"];
  const owners = ["CAD", "Boom", "Production", "Designer", "Manager", "Customer", "Other"];

  const plan = getOrCreateSheet(workbook, planName);
  const examples = getOrCreateSheet(workbook, examplesName);
  const lists = getOrCreateSheet(workbook, listsName);

  lists.setVisibility(ExcelScript.SheetVisibility.visible);

  clearSheet(plan);
  clearSheet(examples);
  clearSheet(lists);

  setupListsSheet(
    lists,
    requestTypes,
    mainCategories,
    subTasks,
    deliverableTypes,
    estimatedTimes,
    priorities,
    statuses,
    owners
  );

  setupTrackerSheet(plan, headers, listsName, false);
  setupTrackerSheet(examples, headers, listsName, true);

  lists.setVisibility(ExcelScript.SheetVisibility.hidden);
  plan.activate();
}

function getOrCreateSheet(workbook: ExcelScript.Workbook, sheetName: string): ExcelScript.Worksheet {
  const existing = workbook.getWorksheet(sheetName);
  if (existing) {
    return existing;
  }
  return workbook.addWorksheet(sheetName);
}

function clearSheet(sheet: ExcelScript.Worksheet) {
  const usedRange = sheet.getUsedRange();
  if (usedRange) {
    usedRange.clear();
  }
}

function setupListsSheet(
  sheet: ExcelScript.Worksheet,
  requestTypes: string[],
  mainCategories: string[],
  subTasks: string[],
  deliverableTypes: string[],
  estimatedTimes: string[],
  priorities: string[],
  statuses: string[],
  owners: string[]
) {
  writeList(sheet, 0, "Request Type", requestTypes);
  writeList(sheet, 1, "Main Category", mainCategories);
  writeList(sheet, 2, "Sub Task", subTasks);
  writeList(sheet, 3, "Deliverable Type", deliverableTypes);
  writeList(sheet, 4, "Estimated Time", estimatedTimes);
  writeList(sheet, 5, "Priority", priorities);
  writeList(sheet, 6, "Status", statuses);
  writeList(sheet, 7, "Owner", owners);

  const headerRange = sheet.getRange("A1:H1");
  headerRange.getFormat().getFill().setColor("#1F4E78");
  headerRange.getFormat().getFont().setColor("#FFFFFF");
  headerRange.getFormat().getFont().setBold(true);
  sheet.getUsedRange()?.getFormat().autofitColumns();
}

function writeList(sheet: ExcelScript.Worksheet, columnIndex: number, header: string, values: string[]) {
  sheet.getCell(0, columnIndex).setValue(header);
  sheet.getRangeByIndexes(1, columnIndex, values.length, 1).setValues(values.map(value => [value]));
}

function setupTrackerSheet(
  sheet: ExcelScript.Worksheet,
  headers: string[],
  listsName: string,
  includeExamples: boolean
) {
  sheet.getRangeByIndexes(0, 0, 1, headers.length).setValues([headers]);
  formatHeader(sheet.getRange("A1:M1"));

  if (includeExamples) {
    const examples = [
      ["2026-W27", "Initial Request", "C0553-F", "CAD", "CAD File", "2 to 3 hours", "High", "In Progress", "CAD", "29 Jun 2026", "30 Jun 2026", "Create CAD file for approved design direction.", ""],
      ["2026-W27", "New Request", "RhinoArtisan", "Test with real job", "Document", "Half day", "Medium", "Waiting", "Boom", "29 Jun 2026", "2 Jul 2026", "Test RhinoArtisan workflow with a real production job.", ""],
      ["2026-W27", "Follow-up", "PD CAD Investigation", "Finding Data", "Document", "This week", "Medium", "In Progress", "Manager", "29 Jun 2026", "5 Jul 2026", "Collect CAD process issues and organize findings.", ""],
      ["2026-W27", "Initial Request", "Cosmic Collection", "Presentation", "Presentation", "1 day", "Low", "Done", "Designer", "28 Jun 2026", "29 Jun 2026", "Prepare CAD example presentation for review.", ""],
      ["2026-W27", "Revision", "Custom Order", "Update Production to CCU and OPC", "Email", "30 to 45 minutes", "High", "Blocked", "Production", "29 Jun 2026", "29 Jun 2026", "Waiting for production confirmation before update.", ""]
    ];
    sheet.getRangeByIndexes(1, 0, examples.length, headers.length).setValues(examples);
  }

  sheet.getFreezePanes().freezeRows(1);
  sheet.getAutoFilter().apply(sheet.getRange("A1:M501"));

  applyDropdown(sheet, "B2:B501", listsName, "A", 4);
  applyDropdown(sheet, "C2:C501", listsName, "B", 14);
  applyDropdown(sheet, "D2:D501", listsName, "C", 15);
  applyDropdown(sheet, "E2:E501", listsName, "D", 12);
  applyDropdown(sheet, "F2:F501", listsName, "E", 9);
  applyDropdown(sheet, "G2:G501", listsName, "F", 3);
  applyDropdown(sheet, "H2:H501", listsName, "G", 5);
  applyDropdown(sheet, "I2:I501", listsName, "H", 7);

  sheet.getRange("J2:K501").setNumberFormatLocal("d mmm yyyy");
  sheet.getRange("L:M").getFormat().setWrapText(true);
  sheet.getRange("A2:M501").getFormat().getFont().setName("Aptos");
  sheet.getRange("A2:M501").getFormat().getFont().setSize(10);
  setColumnWidths(sheet);
}

function applyDropdown(
  sheet: ExcelScript.Worksheet,
  targetAddress: string,
  listsSheetName: string,
  listColumn: string,
  itemCount: number
) {
  const sourceAddress = `='${listsSheetName}'!$${listColumn}$2:$${listColumn}$${itemCount + 1}`;

  sheet.getRange(targetAddress).getDataValidation().setRule({
    list: {
      inCellDropDown: true,
      source: sourceAddress
    }
  });
}

function formatHeader(headerRange: ExcelScript.Range) {
  headerRange.getFormat().getFill().setColor("#1F4E78");
  headerRange.getFormat().getFont().setColor("#FFFFFF");
  headerRange.getFormat().getFont().setBold(true);
  headerRange.getFormat().getFont().setName("Aptos");
  headerRange.getFormat().getFont().setSize(10);
  headerRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  headerRange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  headerRange.getFormat().setRowHeight(28);
}

function setColumnWidths(sheet: ExcelScript.Worksheet) {
  sheet.getRange("A:A").getFormat().setColumnWidth(85);
  sheet.getRange("B:B").getFormat().setColumnWidth(130);
  sheet.getRange("C:C").getFormat().setColumnWidth(150);
  sheet.getRange("D:D").getFormat().setColumnWidth(165);
  sheet.getRange("E:E").getFormat().setColumnWidth(145);
  sheet.getRange("F:F").getFormat().setColumnWidth(125);
  sheet.getRange("G:G").getFormat().setColumnWidth(90);
  sheet.getRange("H:H").getFormat().setColumnWidth(105);
  sheet.getRange("I:I").getFormat().setColumnWidth(105);
  sheet.getRange("J:J").getFormat().setColumnWidth(105);
  sheet.getRange("K:K").getFormat().setColumnWidth(105);
  sheet.getRange("L:L").getFormat().setColumnWidth(260);
  sheet.getRange("M:M").getFormat().setColumnWidth(220);
}
