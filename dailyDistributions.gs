/**
 * Generates an HTML table with multiple columns from a list of agent names.
 * The table will have a maximum of 15 rows, adding new columns as needed.
 *
 * @param {string[]} agents - An array of agent names.
 * @param {number} maxRowsPerColumn - The maximum number of agents to list in a single column.
 * @returns {string} An HTML string representing the formatted table.
 */
function generateAgentTableHtml(agents, maxRowsPerColumn) {
  var totalAgents = agents.length;
  if (totalAgents === 0) {
    return "<p>All active agents have made at least one sale this week.</p>";
  }

  var numColumns = Math.ceil(totalAgents / maxRowsPerColumn);
  
  var html = '<table><thead><tr>';
  for (var c = 0; c < numColumns; c++) {
    html += '<th>Agent Name</th>';
  }
  html += '</tr></thead><tbody>';

  for (var r = 0; r < maxRowsPerColumn; r++) {
    html += '<tr>';
    for (var c = 0; c < numColumns; c++) {
      var index = c * maxRowsPerColumn + r;
      // Use a non-breaking space for empty cells to maintain table structure
      var agentName = index < totalAgents ? agents[index] : ' ';
      html += '<td>' + agentName + '</td>';
    }
    html += '</tr>';
  }

  html += '</tbody></table>';
  return html;
}

/**
 * Sends a daily report summarizing sales for the previous day and listing active agents
 * who have not made any sales in the current week (Monday-Sunday).
 */
function sendDailyReport() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Bi-weekly Stats");
  var sfChqSheet = spreadsheet.getSheetByName("SF + CHQ");
  var lookupSheet = spreadsheet.getSheetByName("Look up Sheet");

  if (!sheet || !sfChqSheet || !lookupSheet) {
    Logger.log("One or more required sheets are missing. Please check sheet names: 'Bi-weekly Stats', 'SF + CHQ', 'Look up Sheet'");
    return;
  }

  // --- 1. Yesterday's Sales Summary ---
  var today = new Date();
  today.setDate(today.getDate() - 1);
  var yesterdayStr = Utilities.formatDate(today, Session.getTimeZone(), "yyyy-MM-dd");

  var data = sheet.getDataRange().getValues();
  var salesSummary = {
    ICS: 0, EPC: 0, WF: 0,
    ICS_OTP: "0.0%", EPC_OTP: "0.0%", WF_OTP: "0.0%"
  };

  for (var i = 2; i < data.length; i++) {
    if (data[i][1] && new Date(data[i][1]).toString() !== "Invalid Date") {
       var dateStr = Utilities.formatDate(new Date(data[i][1]), Session.getTimeZone(), "yyyy-MM-dd");
       if (dateStr === yesterdayStr) {
         salesSummary.ICS += data[i][3] || 0;
         salesSummary.EPC += data[i][4] || 0;
         salesSummary.WF += data[i][5] || 0;
         salesSummary.ICS_OTP = (data[i][6] * 100).toFixed(1) + "%";
         salesSummary.EPC_OTP = (data[i][7] * 100).toFixed(1) + "%";
         salesSummary.WF_OTP = (data[i][8] * 100).toFixed(1) + "%";
       }
    }
  }

  // --- 2. Agents with No Sales in the Current Week ---
  var now = new Date();
  var dayOfWeek = now.getDay();
  var numDaysSinceMonday = (dayOfWeek + 6) % 7;
  
  var startOfWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() - numDaysSinceMonday);
  startOfWeek.setHours(0, 0, 0, 0);

  var endOfWeek = new Date(startOfWeek);
  endOfWeek.setDate(startOfWeek.getDate() + 6);
  endOfWeek.setHours(23, 59, 59, 999);

  var lookupData = lookupSheet.getRange("N2:S").getValues();
  var activeAgents = new Set();
  for (var k = 0; k < lookupData.length; k++) {
    if (!lookupData[k][0]) break;
    if (lookupData[k][5] === "Active") {
      activeAgents.add(lookupData[k][0].trim());
    }
  }

  var sfChqData = sfChqSheet.getRange("A2:F").getValues();
  var agentsWithSalesThisWeek = new Set();
  for (var m = 0; m < sfChqData.length; m++) {
    var agentName = sfChqData[m][0];
    var saleDateValue = sfChqData[m][5];
    
    if (!agentName || !saleDateValue) continue;
    
    var saleDate = new Date(saleDateValue);
    if (saleDate.toString() !== "Invalid Date" && saleDate >= startOfWeek && saleDate <= endOfWeek) {
        agentsWithSalesThisWeek.add(agentName.trim());
    }
  }

  var agentsWithNoSales = [...activeAgents].filter(function(agent) {
    return !agentsWithSalesThisWeek.has(agent);
  });

  // --- 3. Create and Send Styled HTML Email ---
  
  // Generate the multi-column table for agents with no sales
  var noSalesAgentTable = generateAgentTableHtml(agentsWithNoSales, 15);

  var summaryHtml = `
  <html>
  <head>
    <style>
      body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
      h2 { color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 5px; }
      h3 { color: #2c3e50; margin-top: 25px; }
      table { border-collapse: collapse; width: 100%; margin: 15px 0; box-shadow: 0 2px 3px rgba(0,0,0,0.1); table-layout: fixed; }
      th { background-color: #3498db; color: white; text-align: left; padding: 10px; font-weight: bold; }
      td { padding: 8px 10px; border-bottom: 1px solid #ddd; word-wrap: break-word; }
      tr:nth-child(even) { background-color: #f2f2f2; }
      tr:hover { background-color: #e6f7ff; }
      .footer { margin-top: 30px; font-size: 0.9em; color: #7f8c8d; border-top: 1px solid #eee; padding-top: 10px; }
      .highlight { font-weight: bold; color: #e74c3c; }
    </style>
  </head>
  <body>
    <h2>Sales Summary for ${yesterdayStr}</h2>
    <table>
      <tr>
        <th>Product</th>
        <th>Sales</th>
        <th>OTP Verification Rate</th>
      </tr>
      <tr>
        <td>ICS</td>
        <td>${salesSummary.ICS}</td>
        <td>${salesSummary.ICS_OTP}</td>
      </tr>
      <tr>
        <td>EPC</td>
        <td>${salesSummary.EPC}</td>
        <td>${salesSummary.EPC_OTP}</td>
      </tr>
      <tr>
        <td>W.F</td>
        <td>${salesSummary.WF}</td>
        <td>${salesSummary.WF_OTP}</td>
      </tr>
    </table>
    
    <h3>Agents with <span class="highlight">No Sales</span> This Week</h3>
    ${noSalesAgentTable}
    
    <div class="footer">
      <p>Regards,</p>
      <p>Performance Monitoring System</p>
      <p>This is an auto-generated report.</p>
      <p>For any inquiries, contact <a href="mailto:johnpaul@upenergygroup.com">johnpaul@upenergygroup.com</a></p>
    </div>
  </body>
  </html>
  `;

  MailApp.sendEmail({
    to: "distribution.ug@upenergygroup.com",
    subject: "Daily Distribution Summary for " + yesterdayStr,
    htmlBody: summaryHtml,
    noReply: true
  });
}

/**
 * Deletes any existing triggers and creates a new one to run the report daily.
 */
function createDailyTrigger() {
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }

  ScriptApp.newTrigger("sendDailyReport")
    .timeBased()
    .everyDays(1)
    .atHour(10)
    .create();
  
  Logger.log("Trigger created successfully to run 'sendDailyReport' daily at 10 AM.");
}