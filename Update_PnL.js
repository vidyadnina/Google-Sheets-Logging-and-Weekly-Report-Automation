function updateDoneCalculateDaily() {
  const monitoredSheets = ["Margin MMN", "Margin LMK", "Margin VAS"];
  const columnsToMonitor = {
    coaDone: "JG", // COA done? column
    freightInputted: "JH", // Freight inputted? column
    doneCalculate: "JJ", // Done Calculate column
    projectNo: "A", // Project No column
  };

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    monitoredSheets.forEach(sheetName => {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        Logger.log(`Sheet ${sheetName} does not exist.`);
        return;
      }

      const headerRow = 2; // Headers are in row 2
      const lastRow = sheet.getLastRow(); // Get the last non-empty row
      if (lastRow <= headerRow) {
        Logger.log(`No data rows found in sheet ${sheetName}.`);
        return;
      }

      const dataRange = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn());
      const data = dataRange.getValues();
      const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];

      // Identify column indexes
      const projectNoIndex = headers.indexOf("Project No") + 1;
      const coaDoneIndex = headers.indexOf("COA done?") + 1;
      const freightInputtedIndex = headers.indexOf("Freight inputted?") + 1;
      const doneCalculateIndex = headers.indexOf("Done Calculate") + 1;

      if (
        projectNoIndex === 0 ||
        coaDoneIndex === 0 ||
        freightInputtedIndex === 0 ||
        doneCalculateIndex === 0
      ) {
        Logger.log(`Required columns not found in sheet ${sheetName}.`);
        return;
      }

      const updates = []; // Collect rows to update

      for (let i = 0; i < data.length; i++) {
        const projectNo = data[i][projectNoIndex - 1];
        const coaDoneValue = data[i][coaDoneIndex - 1];
        const freightInputtedValue = data[i][freightInputtedIndex - 1];
        const doneCalculateValue = data[i][doneCalculateIndex - 1];

        if (!projectNo) {
          // Skip rows where Project No is empty
          continue;
        }

        if (coaDoneValue === true && freightInputtedValue === true && doneCalculateValue !== true) {
          updates.push([i + headerRow + 1, true]); // Collect row index and value
        } else if (doneCalculateValue !== "" && (coaDoneValue !== true || freightInputtedValue !== true)) {
          updates.push([i + headerRow + 1, ""]); // Collect row index to clear
        }
      }

      // Apply updates in bulk
      if (updates.length > 0) {
        const doneCalculateColumnIndex = doneCalculateIndex;
        updates.forEach(([rowIndex, value]) => {
          sheet.getRange(rowIndex, doneCalculateColumnIndex).setValue(value);
        });
        Logger.log(`Updated ${updates.length} rows in sheet ${sheetName}.`);
      } else {
        Logger.log(`No updates required for sheet ${sheetName}.`);
      }
    });
  } catch (error) {
    Logger.log(`Error in updateDoneCalculateWeekly: ${error}`);
  }
}
