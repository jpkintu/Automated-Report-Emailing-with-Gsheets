function calculateTargetUpdatesAndSendEmail() {
  // Get yesterday's date to determine the month (1-12)
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const currentMonth = yesterday.getMonth() + 1;
  const monthNames = ["January", "February", "March", "April", "May", "June", 
                     "July", "August", "September", "October", "November", "December"];
  
  // Open spreadsheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Bi-Weekly Stats");
  
  // Get all data (headers in row 2)
  const data = sheet.getDataRange().getValues();
  
  // Column indices (0-based)
  const monthCol = 10; // Column K
  const icsTargetCol = 11; // Column L
  const icsActualCol = 12; // Column M
  const wfTargetCol = 13; // Column N
  const wfActualCol = 14; // Column O
  const epcTargetCol = 15; // Column P
  const epcActualCol = 16; // Column Q
  const agentMonthCol = 37; // Column AL
  const icsAgentCol = 38; // Column AM
  const wfAgentCol = 40; // Column AO
  const epcAgentCol = 39; // Column AN
  
  // OTP column indices (0-based, starting from row 8)
  const otpMonthCol = 17; // Column R
  const icsOtpCol = 18; // Column S
  const wfOtpCol = 19; // Column T
  const epcOtpCol = 20; // Column U
  
  // Filter data for current month (skip header rows)
  const currentMonthData = data.filter((row, index) => {
    if (index < 2) return false; // Skip header rows
    
    // Handle both numeric and date month values
    const monthValue = row[monthCol];
    let monthNum;
    
    if (monthValue instanceof Date) {
      monthNum = monthValue.getMonth() + 1;
    } else {
      monthNum = Number(monthValue);
    }
    
    return monthNum === currentMonth;
  });
  
  // Get agent counts for current month (from first row that matches)
  let icsAgents = 0;
  let wfAgents = 0;
  let epcAgents = 0;
  
  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    const monthValue = row[monthCol];
    let monthNum;
    
    if (monthValue instanceof Date) {
      monthNum = monthValue.getMonth() + 1;
    } else {
      monthNum = Number(monthValue);
    }
    
    if (monthNum === currentMonth) {
      icsAgents = Number(row[icsAgentCol]) || 0;
      wfAgents = Number(row[wfAgentCol]) || 0;
      epcAgents = Number(row[epcAgentCol]) || 0;
      break; // We just need the first matching row
    }
  }
  
  // Get OTP values for current month (starting from row 8)
  let icsOtp = 0;
  let wfOtp = 0;
  let epcOtp = 0;
  
  // Get OTP data range (starting from row 8)
  const otpDataRange = sheet.getRange(8, otpMonthCol + 1, sheet.getLastRow() - 7, 4); // Columns R-U
  const otpData = otpDataRange.getValues();
  
  for (let i = 0; i < otpData.length; i++) {
    const row = otpData[i];
    const monthValue = row[0]; // Column R (0-based index 0)
    let monthNum;
    
    if (monthValue instanceof Date) {
      monthNum = monthValue.getMonth() + 1;
    } else {
      monthNum = Number(monthValue);
    }
    
    if (monthNum === currentMonth) {
      icsOtp = Number(row[icsOtpCol - otpMonthCol]) || 0; // Column S (index 1)
      wfOtp = Number(row[wfOtpCol - otpMonthCol]) || 0; // Column T (index 2)
      epcOtp = Number(row[epcOtpCol - otpMonthCol]) || 0; // Column U (index 3)
      break;
    }
  }
  
  // Calculate metrics for each product type
  const products = [
    { name: "ICS", targetCol: icsTargetCol, actualCol: icsActualCol, color: "#e6f3ff", agents: icsAgents, otp: icsOtp },
    { name: "W.F", targetCol: wfTargetCol, actualCol: wfActualCol, color: "#fff2e6", agents: wfAgents, otp: wfOtp },
    { name: "EPC", targetCol: epcTargetCol, actualCol: epcActualCol, color: "#e6ffe6", agents: epcAgents, otp: epcOtp }
  ];
  
  let summary = [];
  let efficiencySummary = [];
  let otpSummary = [];
  const today = new Date();
  const yesterdayDate = yesterday.getDate();
  const daysInMonth = new Date(yesterday.getFullYear(), yesterday.getMonth() + 1, 0).getDate();
  const daysElapsed = yesterdayDate - 1;
  const daysRemaining = daysInMonth - yesterdayDate + 1;
  
  products.forEach(product => {
    const target = currentMonthData.reduce((sum, row) => sum + (Number(row[product.targetCol])) || 0, 0);
    const actual = currentMonthData.reduce((sum, row) => sum + (Number(row[product.actualCol])) || 0, 0);
    const difference = target - actual;
    
    // Calculate percentages
    const percentageAchieved = target > 0 ? (actual / target) * 100 : 0;
    const percentageExpected = (yesterdayDate / daysInMonth) * 100;
    const percentageShortfall = percentageExpected - percentageAchieved;
    
    // Calculate daily requirements
    const remainingShortfall = Math.max(0, difference);
    const dailyRequired = remainingShortfall > 0 && daysRemaining > 0 ? 
                         (remainingShortfall / daysRemaining) : 0;
    
    // Calculate agent efficiency
    const efficiency = product.agents > 0 ? (actual / product.agents) : 0;
    
    summary.push({
      Product: product.name,
      Target: target,
      Achieved: actual,
      Difference: difference,
      PercentageAchieved: percentageAchieved.toFixed(2) + "%",
      PercentageExpected: percentageExpected.toFixed(2) + "%",
      PercentageShortfall: percentageShortfall.toFixed(2) + "%",
      DailyRequired: dailyRequired.toFixed(2),
      DaysElapsed: daysElapsed,
      DaysRemaining: daysRemaining,
      color: product.color
    });
    
    efficiencySummary.push({
      Product: product.name,
      Agents: product.agents,
      TotalEfficiency: efficiency.toFixed(2),
      color: product.color
    });
    
    otpSummary.push({
      Product: product.name,
      OTP: product.otp.toFixed(2) + "%",
      color: product.color
    });
  });
  
  // Generate email content with styling
  let emailBody = `
  <html>
  <head>
    <style>
      body {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        line-height: 1.6;
        color: #333;
        max-width: 800px;
        margin: 0 auto;
        padding: 20px;
      }
      h2 {
        color: #2c3e50;
        border-bottom: 2px solid #3498db;
        padding-bottom: 8px;
        margin-bottom: 20px;
      }
      .info-box {
        background-color: #f8f9fa;
        border-left: 4px solid #3498db;
        padding: 12px;
        margin-bottom: 20px;
      }
      table {
        border-collapse: collapse;
        width: 100%;
        margin: 20px 0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
      }
      th {
        background-color: #3498db;
        color: white;
        text-align: center;
        padding: 12px;
        font-weight: 600;
      }
      td {
        padding: 10px;
        border-bottom: 1px solid #e0e0e0;
        text-align: center;
      }
      tr:nth-child(even) {
        background-color: #f9f9f9;
      }
      .positive {
        color: #27ae60;
        font-weight: bold;
      }
      .negative {
        color: #e74c3c;
        font-weight: bold;
      }
      .key {
        margin: 20px 0;
        padding: 15px;
        background-color: #f8f9fa;
        border-radius: 4px;
      }
      .key li {
        margin-bottom: 8px;
      }
      .footer {
        margin-top: 30px;
        font-size: 0.9em;
        color: #7f8c8d;
        border-top: 1px solid #eee;
        padding-top: 15px;
      }
      .section-title {
        margin-top: 30px;
        color: #2c3e50;
        font-size: 1.2em;
        font-weight: 600;
      }
      
      /* Side by side tables styling */
      .row {
        margin-left:-5px;
        margin-right:-5px;
        display: flex;
      }
        
      .column {
        float: left;
        width: 50%;
        padding: 5px;
      }

      /* Clearfix (clear floats) */
      .row::after {
        content: "";
        clear: both;
        display: table;
      }
    </style>
  </head>
  <body>
    <h2>Monthly Performance Summary - ${monthNames[currentMonth-1]}</h2>
    
    <div class="info-box">
      <p><strong>Reporting Date:</strong> ${new Date().toDateString()}</p>
      <p><strong>Month Progress:</strong> ${yesterdayDate}/${daysInMonth} days (${((yesterdayDate/daysInMonth)*100).toFixed(1)}%)</p>
    </div>
    
    <div class="section-title">Performance Metrics</div>
    <table>
      <tr>
        <th>Product</th>
        <th>Target</th>
        <th>Achieved</th>
        <th>Difference</th>
        <th>% Achieved</th>
        <th>% Expected</th>
        <th>% Shortfall</th>
        <th>Daily Required</th>
      </tr>
  `;
  
  summary.forEach(item => {
    const isOnTrack = item.Difference <= 0;
    const isShortfallPositive = Number(item.PercentageShortfall.replace('%','')) > 0;
    
    emailBody += `
      <tr style="background-color: ${item.color}">
        <td style="font-weight: bold;">${item.Product}</td>
        <td>${item.Target.toLocaleString()}</td>
        <td>${item.Achieved.toLocaleString()}</td>
        <td class="${isOnTrack ? 'positive' : 'negative'}">
          ${isOnTrack ? '+' : ''}${item.Difference.toLocaleString()}
        </td>
        <td class="${isOnTrack ? 'positive' : 'negative'}">${item.PercentageAchieved}</td>
        <td>${item.PercentageExpected}</td>
        <td class="${isShortfallPositive ? 'negative' : 'positive'}">${item.PercentageShortfall}</td>
        <td>${item.DailyRequired}</td>
      </tr>
    `;
  });
   
  emailBody += `</table>
  
    <div class="section-title">Agent Efficiency & OTP Metrics</div>
    <div class="row">
      <div class="column">
        <table>
          <tr>
            <th>Product</th>
            <th>Number of Agents</th>
            <th>Efficiency</th>
          </tr>
  `;
  
  efficiencySummary.forEach(item => {
    emailBody += `
      <tr style="background-color: ${item.color}">
        <td style="font-weight: bold;">${item.Product}</td>
        <td>${item.Agents}</td>
        <td>${item.TotalEfficiency}</td>
      </tr>
    `;
  });
  
  emailBody += `</table>
      </div>
      <div class="column">
        <table>
          <tr>
            <th>Product</th>
            <th>OTP Percentage</th>
          </tr>
  `;
  
  otpSummary.forEach(item => {
    const otpValue = parseFloat(item.OTP);
    
    emailBody += `
      <tr style="background-color: ${item.color}">
        <td style="font-weight: bold;">${item.Product}</td>
        <td>${item.OTP}</td>
      </tr>
    `;
  }); 

  emailBody += `</table>
      </div>
    </div>
    
    <div class="key">
      <p><strong>Key:</strong></p>
      <ul>
        <li><strong>Difference:</strong> Target - Achieved (positive = behind, negative = ahead)</li>
        <li><strong>% Achieved:</strong> Current performance vs total target</li>
        <li><strong>% Expected:</strong> Where we should be based on calendar progress</li>
        <li><strong>% Shortfall:</strong> Expected % minus Achieved % (positive = behind, negative = ahead)</li>
        <li><strong>Daily Required:</strong> Amount needed per remaining day to hit target</li>
        <li><strong>Efficiency:</strong> Total achieved divided by number of agents</li>
        </li>
      </ul>
    </div>
    
    <div class="footer" style="text-align: left; margin-top: 15px; border-top: 1px solid #eee; padding-top: 15px;">
  <p>Regards,</p>
  <p>Performance Monitoring System</p>
  <p><em>Note:This is an auto-generated report. For any inquiries, contact <a href="mailto:jp@example.com">The Distribution Team</a></em></p>
  
  <div style="margin-top: 15px; text-align: left;">
    <img src="images.squarespace-cdn.com/content/v1/60a3e4977dbfad2879614ea3/15785e2e-4439-4d46-ae1f-75ff8c7940fe/logo.png?format=1500w" 
         alt="UpEnergy Logo" 
         style="max-width: 200px; height: auto;">
  </div>
</div>
  </body>
  </html>
  `;
  
  // Send email
  const recipients = "jp@example.com";
  const subject = `Distribution Performance Update - ${monthNames[currentMonth-1]} (Day ${yesterdayDate}/${daysInMonth})`;
  
  MailApp.sendEmail({
    to: recipients,
    subject: subject,
    htmlBody: emailBody,
    noReply: true
  });
  
  console.log("Email sent successfully!");
}
