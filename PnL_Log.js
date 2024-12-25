function logMarginUpdates() {
  const monitoredSheets = ["Margin MMN", "Margin LMK", "Margin VAS"];
  const logSheetId = "1hcdQNVOeKGJxbxB5s00fvvs_woOCPfYIhols3ZdPR7o"; // Log spreadsheet ID
  const logSheetName = "Log Margin"; // Log sheet name
  const stateSheetName = "Log States"; // Hidden sheet to store previous states
  const columnsToMonitor = {
    doneCalculate: "JJ", // Done Calculate column
    projectNo: "A", // Project No column
    route: "G", // Route column
  };

  try {
    const logSpreadsheet = SpreadsheetApp.openById(logSheetId);
    let logSheet = logSpreadsheet.getSheetByName(logSheetName);
    let stateSheet = logSpreadsheet.getSheetByName(stateSheetName);

    // If the log sheet doesn't exist, create it
    if (!logSheet) {
      logSheet = logSpreadsheet.insertSheet(logSheetName);
      logSheet.appendRow(["Timestamp", "PT", "Project No", "Route", "Week Group"]); // Add headers
    }

    // If the state sheet doesn't exist, create it
    if (!stateSheet) {
      stateSheet = logSpreadsheet.insertSheet(stateSheetName);
      stateSheet.appendRow(["Project No", "PT", "Done Calculate"]); // Add headers
      stateSheet.hideSheet(); // Hide the sheet to keep it clean
    }

    const stateData = stateSheet.getDataRange().getValues();
    const stateHeaders = stateData.shift(); // Remove headers
    const stateMap = new Map(
      stateData.map(row => [`${row[1]}|${row[0]}`, row[2]]) // Create a map: "PT|Project No" => Done Calculate State
    );

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const now = new Date();
    const weekGroup = getWeekGroup(now); // Get the current week group
    const logsToAppend = [];
    const newStateData = [];

    monitoredSheets.forEach(sheetName => {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        Logger.log(`Sheet ${sheetName} does not exist.`);
        return;
      }

      const headerRow = 2; // Header is in row 2
      const dataRange = sheet.getDataRange();
      const data = dataRange.getValues();
      const headers = data[headerRow - 1]; // Get headers (row 2)

      // Identify column indexes
      const doneCalculateIndex = headers.indexOf("Done Calculate") + 1;
      const projectNoIndex = headers.indexOf("Project No") + 1;
      const routeIndex = headers.indexOf("Route") + 1;

      if (
        doneCalculateIndex === 0 ||
        projectNoIndex === 0 ||
        routeIndex === 0
      ) {
        Logger.log(`Required columns not found in sheet ${sheetName}.`);
        return;
      }

      // Loop through each row after the header
      for (let i = headerRow; i < data.length; i++) {
        const doneCalculateValue = data[i][doneCalculateIndex - 1]; // Done Calculate value
        const projectNo = data[i][projectNoIndex - 1];
        const route = data[i][routeIndex - 1];
        const pt = sheetName.split(" ")[1]; // Extract PT (MMN, LMK, VAS) from sheet name
        const uniqueKey = `${pt}|${projectNo}`;

        if (!projectNo) {
          // Skip rows where Project No is empty
          continue;
        }

        // Check previous state
        const previousState = stateMap.get(uniqueKey) || ""; // Default to empty if not found

        // Log only if Done Calculate transitioned from empty to TRUE
        if (doneCalculateValue === true && previousState !== true) {
          Logger.log(`Logging update for Project No: ${projectNo} in sheet ${sheetName}`);
          logsToAppend.push([new Date(), pt, projectNo, route, weekGroup]);
        }

        // Update state map with the current state
        newStateData.push([projectNo, pt, doneCalculateValue]);
      }
    });

    // Append all collected logs to the log sheet at once
    if (logsToAppend.length > 0) {
      logSheet.getRange(logSheet.getLastRow() + 1, 1, logsToAppend.length, logsToAppend[0].length).setValues(logsToAppend);
      Logger.log(`Logged ${logsToAppend.length} updates to "${logSheetName}".`);
    } else {
      Logger.log("No new updates to log.");
    }

    // Update state sheet with new state data
    stateSheet.clear();
    stateSheet.appendRow(stateHeaders); // Re-add headers
    stateSheet.getRange(2, 1, newStateData.length, newStateData[0].length).setValues(newStateData);
  } catch (error) {
    Logger.log(`Error in logMarginUpdates: ${error}`);
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
