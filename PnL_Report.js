function sendMarginSummary() {
  const logSheetId = "1hcdQNVOeKGJxbxB5s00fvvs_woOCPfYIhols3ZdPR7o"; // Log spreadsheet ID
  const logSheetName = "Log Margin"; // Log sheet name
  const emailRecipient = "vidya.liken@lieco.co.id"; // Replace with your email address

  try {
    const logSpreadsheet = SpreadsheetApp.openById(logSheetId);
    const logSheet = logSpreadsheet.getSheetByName(logSheetName);

    if (!logSheet) {
      Logger.log(`Log sheet "${logSheetName}" does not exist.`);
      return;
    }

    const now = new Date();
    const lastMonday = new Date(now.setDate(now.getDate() - now.getDay() + 1 - 7)); // Last Monday
    const today = new Date(); // Today

    const data = logSheet.getDataRange().getValues();
    const headers = data.shift(); // Remove header row
    const recentLogs = data.filter(row => {
      const logDate = new Date(row[0]); // Timestamp
      return logDate >= lastMonday && logDate <= today; // Filter up to today
    });

    if (recentLogs.length === 0) {
      Logger.log("No updates to notify.");
      return;
    }

    // Format the email
    const subject = `Weekly Margin Update Summary - ${getWeekGroup(today)}`; // Use current week
    let htmlTable = `<table border="1" style="border-collapse: collapse; width: 100%;">`;
    htmlTable += `<tr style="background-color: #f2f2f2;">
                    <th style="padding: 8px; text-align: left;">Project No</th>
                    <th style="padding: 8px; text-align: left;">PT</th>
                    <th style="padding: 8px; text-align: left;">Route</th>
                    <th style="padding: 8px; text-align: left;">Week Group</th>
                  </tr>`;
    recentLogs.forEach(row => {
      htmlTable += `<tr>
                      <td style="padding: 8px; text-align: left;">${row[2]}</td>
                      <td style="padding: 8px; text-align: left;">${row[1]}</td>
                      <td style="padding: 8px; text-align: left;">${row[3]}</td>
                      <td style="padding: 8px; text-align: left;">${row[4]}</td>
                    </tr>`;
    });
    htmlTable += `</table>`;

    const body = `
      <p>Dear Team,</p>
      <p>Here is the list of projects marked as "Done Calculated" for the past week:</p>
      ${htmlTable}
      <p>Best regards,<br>Your Automated Logging System</p>
    `;

    MailApp.sendEmail({
      to: emailRecipient,
      subject: subject,
      htmlBody: body,
    });

    Logger.log(`Email sent to ${emailRecipient} with ${recentLogs.length} updates.`);
  } catch (error) {
    Logger.log(`Error in sendMarginSummary: ${error}`);
  }
}
