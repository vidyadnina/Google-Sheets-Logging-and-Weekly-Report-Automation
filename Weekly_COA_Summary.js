function sendWeeklySummary() {
  const logSheetId = "1hcdQNVOeKGJxbxB5s00fvvs_woOCPfYIhols3ZdPR7o"; // Log sheet ID
  const emailToNotify = "vidya.liken@lieco.co.id"; // Email address to notify
  
  const logSheet = SpreadsheetApp.openById(logSheetId).getSheetByName("Log COA");
  if (!logSheet) {
    Logger.log("Log sheet does not exist.");
    return;
  }
  
  const now = new Date();
  const lastWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 7);
  
  const data = logSheet.getDataRange().getValues();
  const headers = data.shift(); // Remove header row
  
  // Filter rows updated in the last week
  const recentUpdates = data.filter(row => new Date(row[0]) >= lastWeek);
  
  if (recentUpdates.length === 0) {
    MailApp.sendEmail(emailToNotify, "Weekly Update Summary", "No updates were logged last week.");
    return;
  }

  // Build the HTML table
  let htmlTable = `<table border="1" style="border-collapse: collapse; width: 100%;">`;
  htmlTable += `<tr style="background-color: #f2f2f2;">`;
  headers.slice(1).forEach(header => { // Skip the "Timestamp" column
    htmlTable += `<th style="padding: 8px; text-align: left;">${header}</th>`;
  });
  htmlTable += `</tr>`;
  recentUpdates.forEach(row => {
    htmlTable += `<tr>`;
    row.slice(1).forEach(cell => { // Skip the "Timestamp" column
      htmlTable += `<td style="padding: 8px; text-align: left;">${cell || ''}</td>`;
    });
    htmlTable += `</tr>`;
  });
  htmlTable += `</table>`;

  // Determine the current week group
  const currentWeekGroup = getWeekGroup(now);

  // Send the email
  const subject = `COA Updates - ${currentWeekGroup}`;
  const body = `
    <p>Here is the summary of updates logged in the past week:</p>
    ${htmlTable}
    <p>Sekian dan terima gaji,<br>Vidya</p>
  `;
  MailApp.sendEmail({
    to: emailToNotify,
    subject: subject,
    htmlBody: body,
  });

  Logger.log("Weekly summary email sent successfully.");
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
