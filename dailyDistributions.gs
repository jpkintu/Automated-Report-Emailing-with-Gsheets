function sendBiweeklyStats() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bi-weekly Stats");
  var sfSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SF Looker Update");
  
  var today = new Date();
  today.setDate(today.getDate() - 1); // Get yesterday's date
  var yesterdayStr = Utilities.formatDate(today, Session.getTimeZone(), "yyyy-MM-dd");

  // Fetch Data
  var data = sheet.getDataRange().getValues();
  var sfData = sfSheet.getDataRange().getValues();

  var salesSummary = {
    ICS: 0, EPC: 0, WF: 0,
    ICS_OTP: "", EPC_OTP: "", WF_OTP: ""
  };

  var agentSales = {};

  for (var i = 2; i < data.length; i++) {
    var dateStr = Utilities.formatDate(new Date(data[i][1]), Session.getTimeZone(), "yyyy-MM-dd");
    if (dateStr === yesterdayStr) {
      salesSummary.ICS += data[i][3];
      salesSummary.EPC += data[i][4];
      salesSummary.WF += data[i][5];
      salesSummary.ICS_OTP = (data[i][6]*100).toFixed(1) + "%";
      salesSummary.EPC_OTP = (data[i][7]*100).toFixed(1) + "%";
      salesSummary.WF_OTP = (data[i][8]*100).toFixed(1) + "%";
    }
  }

  for (var j = 1; j < sfData.length; j++) {
    var sfDateStr = Utilities.formatDate(new Date(sfData[j][0]), Session.getTimeZone(), "yyyy-MM-dd");
    if (sfDateStr === yesterdayStr && sfData[j][2] === "ICS") {
      var agent = sfData[j][1];
      agentSales[agent] = (agentSales[agent] || 0) + sfData[j][4];
    }
  }

  var underperformingAgents = Object.keys(agentSales).filter(agent => agentSales[agent] < 20);

  // Create styled HTML email
  var summaryHtml = `
  <html>
  <head>
    <style>
      body {
        font-family: Arial, sans-serif;
        line-height: 1.6;
        color: #333;
      }
      h2 {
        color: #2c3e50;
        border-bottom: 2px solid #3498db;
        padding-bottom: 5px;
      }
      h3 {
        color: #2c3e50;
        margin-top: 20px;
      }
      table {
        border-collapse: collapse;
        width: 100%;
        margin: 15px 0;
        box-shadow: 0 2px 3px rgba(0,0,0,0.1);
      }
      th {
        background-color: #3498db;
        color: white;
        text-align: left;
        padding: 10px;
      }
      td {
        padding: 8px 10px;
        border-bottom: 1px solid #ddd;
      }
      tr:nth-child(even) {
        background-color: #f2f2f2;
      }
      tr:hover {
        background-color: #e6f7ff;
      }
      .footer {
        margin-top: 30px;
        font-size: 0.9em;
        color: #7f8c8d;
        border-top: 1px solid #eee;
        padding-top: 10px;
      }
      .highlight {
        font-weight: bold;
        color: #e74c3c;
      }
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
    
    <h3>Agents with <span class="highlight">&lt;20 ICS Sales</span></h3>
    <table>
      <tr>
        <th>Agent Name</th>
        <th>ICS Sales</th>
      </tr>
      ${underperformingAgents.map(agent => `
        <tr>
          <td>${agent}</td>
          <td>${agentSales[agent]}</td>
        </tr>
      `).join('')}
    </table>
    
    <div class="footer">
      <p>Regards,</p>
      <p>Performance Monitoring System</p>
      <p>This is an auto-generated report.</p>
      <p>For any inquiries, contact <a href="mailto:jp@example.com">jp@example.com</a></p>
    </div>
  </body>
  </html>
  `;

  // Send email
  MailApp.sendEmail({
    to: "jp@example.com",
    subject: "Daily Distribution Summary for " + yesterdayStr,
    htmlBody: summaryHtml,
    noReply: true
  });
}

// Set a trigger to run daily at 10 AM
function createTrigger() {
  ScriptApp.newTrigger("sendBiweeklyStats")
    .timeBased()
    .everyDays(1)
    .atHour(10)
    .create();
}
