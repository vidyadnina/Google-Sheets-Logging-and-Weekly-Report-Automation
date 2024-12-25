function onEdit(e) {
  const sheetName = "Shipment Schedule";
  const logSheetId = "1hcdQNVOeKGJxbxB5s00fvvs_woOCPfYIhols3ZdPR7o"; // Log sheet ID (separate spreadsheet)
  const lpGroup = ["Ni LP", "MC LP"]; // LP group columns
  const dpGroup = ["Ni DP", "MC DP"]; // DP group columns

  try {
    Logger.log("onEdit triggered");

    if (!e) {
      Logger.log("No event object passed to onEdit.");
      return;
    }

    const sheet = e.source.getActiveSheet();
    Logger.log(`Edited sheet: ${sheet.getName()}`);

    if (sheet.getName() !== sheetName) {
      Logger.log(`Edited sheet is not ${sheetName}, no action taken.`);
      return;
    }

    const range = e.range;
    const editedColumn = range.getColumn();
    const editedRow = range.getRow();
    Logger.log(`Edited row: ${editedRow}, Edited column index: ${editedColumn}`);

    const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    const columnName = headers[editedColumn - 1];
    Logger.log(`Edited column name: ${columnName}`);

    if (![...lpGroup, ...dpGroup].includes(columnName)) {
      Logger.log(`Edited column ${columnName} is not part of LP or DP groups.`);
      return;
    }

    const projectNo = sheet.getRange(editedRow, headers.indexOf("Project No") + 1).getValue();
    const blDate = sheet.getRange(editedRow, headers.indexOf("BL Date") + 1).getValue();
    const source = sheet.getRange(editedRow, headers.indexOf("Source") + 1).getValue();
    const enduser = sheet.getRange(editedRow, headers.indexOf("Enduser") + 1).getValue();
    Logger.log(`Project No: ${projectNo}, BL Date: ${blDate}, Source: ${source}, Enduser: ${enduser}`);

    const updatedValue = range.getValue();
    Logger.log(`Updated value: ${updatedValue}`);

    const logSpreadsheet = SpreadsheetApp.openById(logSheetId);
    let logSheet = logSpreadsheet.getSheetByName("Log COA");

    if (!logSheet) {
      Logger.log("Log COA sheet not found. Creating a new one...");
      logSheet = logSpreadsheet.insertSheet("Log COA");
      logSheet.appendRow(["Timestamp", "BL Date", "Project No", "Source", "Enduser", "Column Updated", "LP Done", "DP Done", "Week Group"]); // Add headers
    }

    const lpDone =
      sheet.getRange(editedRow, headers.indexOf("Ni LP") + 1).getValue() &&
      sheet.getRange(editedRow, headers.indexOf("MC LP") + 1).getValue()
        ? "Yes"
        : "No";

    const dpDone =
      sheet.getRange(editedRow, headers.indexOf("Ni DP") + 1).getValue() &&
      sheet.getRange(editedRow, headers.indexOf("MC DP") + 1).getValue()
        ? "Yes"
        : "No";

    let columnUpdated = "";
    if (lpGroup.includes(columnName) && lpDone === "Yes") {
      columnUpdated = "LP";
    } else if (dpGroup.includes(columnName) && dpDone === "Yes") {
      columnUpdated = "DP";
    } else {
      Logger.log("No complete group update (LP or DP), skipping log entry.");
      return;
    }
    Logger.log(`Column Updated: ${columnUpdated}`);

    const timestamp = new Date();
    const weekGroup = getWeekGroup(timestamp);
    Logger.log(`Week Group: ${weekGroup}`);

    logSheet.appendRow([timestamp, blDate, projectNo, source, enduser, columnUpdated, lpDone, dpDone, weekGroup]);
    Logger.log(`Logged update for Project No: ${projectNo}, Column: ${columnUpdated}`);
  } catch (error) {
    Logger.log(`Error in onEdit: ${error}`);
  }
}

function getWeekGroup(date) {
  const monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];

  const month = monthNames[date.getMonth()];
  const firstDayOfMonth = new Date(date.getFullYear(), date.getMonth(), 1);
  const dayOfWeek = firstDayOfMonth.getDay(); // Day of the week for the 1st of the month
  const offsetDays = dayOfWeek > 0 ? 7 - dayOfWeek : 0; // Offset to next Sunday
  const firstSunday = new Date(date.getFullYear(), date.getMonth(), 1 + offsetDays);

  const weekOfMonth = Math.ceil((date.getDate() - firstSunday.getDate() + 1) / 7) + (date < firstSunday ? 0 : 1);

  const weekSuffix = getOrdinalSuffix(weekOfMonth); // Get ordinal suffix (1st, 2nd, etc.)
  return `${month} ${weekOfMonth}${weekSuffix} week`;
}

function getOrdinalSuffix(number) {
  if (number === 1) return "st";
  if (number === 2) return "nd";
  if (number === 3) return "rd";
  return "th";
}
