/**
 * Party Quiz Leaderboard Generator
 * 
 * This script creates a beautifully formatted leaderboard from team data.
 * Sorting rules:
 * 1. Higher total points = better ranking
 * 2. If points are equal, fewer people = better ranking
 * 
 * HOW TO INSTALL:
 * 1. Open your Google Spreadsheet
 * 2. Go to Extensions → Apps Script
 * 3. Delete any existing code
 * 4. Paste this entire script
 * 5. Save (Ctrl+S) and name it "Leaderboard"
 * 6. Close the script editor
 * 7. Refresh your spreadsheet
 * 8. You'll see a new menu "🏆 Leaderboard" in the menu bar
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

const CONFIG = {
  // Source sheet settings
  SOURCE_SHEET: 'vysledky',  // Name of sheet with team data
  NAME_COLUMN: 'A',           // Column with team names
  TOTAL_COLUMN: 'B',          // Column with total points
  COUNT_COLUMN: 'C',          // Column with number of people
  START_ROW: 2,               // First data row (skip header)
  
  // Output sheet settings
  OUTPUT_SHEET: 'Leaderboard',
  
  // Visual settings
  COLORS: {
    gold: '#FFD700',      // 1st place
    silver: '#C0C0C0',    // 2nd place
    bronze: '#CD7F32',    // 3rd place
    header: '#1a73e8',    // Header background
    headerText: '#ffffff', // Header text
    evenRow: '#f3f3f3',   // Even row background
    oddRow: '#ffffff',    // Odd row background
    border: '#cccccc'     // Border color
  },
  
  // Trophy emojis
  TROPHIES: {
    1: '🥇',
    2: '🥈',
    3: '🥉'
  }
};

// ============================================================================
// MAIN FUNCTIONS
// ============================================================================

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏆 Leaderboard')
    .addItem('📊 Generate Leaderboard', 'generateLeaderboard')
    .addItem('🔄 Refresh Leaderboard', 'generateLeaderboard')
    .addSeparator()
    .addItem('🌐 View as HTML Page', 'showHtmlLeaderboard')
    .addItem('📱 Get HTML Page URL', 'getWebAppUrl')
    .addSeparator()
    .addItem('⚙️ Configure Settings', 'showConfigDialog')
    .addToUi();
}

/**
 * Main function to generate the leaderboard
 */
function generateLeaderboard() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get source data
    const sourceSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);
    if (!sourceSheet) {
      showError(`Source sheet "${CONFIG.SOURCE_SHEET}" not found!`);
      return;
    }
    
    const data = readTeamData(sourceSheet);
    
    if (data.length === 0) {
      showError('No data found in source sheet!');
      return;
    }
    
    // Sort data
    const sortedData = sortTeams(data);
    
    // Create or clear output sheet
    let outputSheet = ss.getSheetByName(CONFIG.OUTPUT_SHEET);
    if (outputSheet) {
      outputSheet.clear();
    } else {
      outputSheet = ss.insertSheet(CONFIG.OUTPUT_SHEET);
    }
    
    // Write and format leaderboard
    writeLeaderboard(outputSheet, sortedData);
    
    // Activate the leaderboard sheet
    outputSheet.activate();
    
    SpreadsheetApp.getUi().alert(
      '✅ Success!',
      `Leaderboard created with ${sortedData.length} teams!`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    showError('Error generating leaderboard: ' + error.message);
  }
}

/**
 * Reads team data from source sheet
 */
function readTeamData(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.START_ROW) {
    return [];
  }
  
  const data = [];
  
  for (let i = CONFIG.START_ROW; i <= lastRow; i++) {
    const name = sheet.getRange(CONFIG.NAME_COLUMN + i).getValue();
    const total = sheet.getRange(CONFIG.TOTAL_COLUMN + i).getValue();
    const count = sheet.getRange(CONFIG.COUNT_COLUMN + i).getValue();
    
    // Skip empty rows
    if (name === '' && total === '' && count === '') {
      continue;
    }
    
    data.push({
      name: String(name),
      total: Number(total) || 0,
      count: Number(count) || 0
    });
  }
  
  return data;
}

/**
 * Sorts teams according to rules:
 * 1. Higher total = better
 * 2. If equal total, lower count = better
 */
function sortTeams(data) {
  return data.sort((a, b) => {
    // First compare by total (descending)
    if (b.total !== a.total) {
      return b.total - a.total;
    }
    // If totals are equal, compare by count (ascending - fewer is better)
    return a.count - b.count;
  });
}

/**
 * Writes and formats the leaderboard
 */
function writeLeaderboard(sheet, data) {
  // Set column widths
  sheet.setColumnWidth(1, 80);   // Rank
  sheet.setColumnWidth(2, 300);  // Team Name
  sheet.setColumnWidth(3, 120);  // Points
  sheet.setColumnWidth(4, 100);  // People
  
  // Write header
  const headers = [['#', 'Tým', 'Body', 'Počet lidí']];
  const headerRange = sheet.getRange(1, 1, 1, 4);
  headerRange.setValues(headers);
  
  // Format header
  headerRange
    .setBackground(CONFIG.COLORS.header)
    .setFontColor(CONFIG.COLORS.headerText)
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  // Set header row height
  sheet.setRowHeight(1, 40);
  
  // Write data
  for (let i = 0; i < data.length; i++) {
    const row = i + 2; // Start from row 2 (after header)
    const rank = i + 1;
    const team = data[i];
    
    // Prepare row data
    const trophy = CONFIG.TROPHIES[rank] || '';
    const rankDisplay = trophy ? `${trophy} ${rank}` : rank;
    
    const rowData = [[
      rankDisplay,
      team.name,
      team.total,
      team.count
    ]];
    
    // Write row
    const rowRange = sheet.getRange(row, 1, 1, 4);
    rowRange.setValues(rowData);
    
    // Format row
    formatRow(sheet, row, rank, data.length);
  }
  
  // Add borders to entire table
  const tableRange = sheet.getRange(1, 1, data.length + 1, 4);
  tableRange.setBorder(
    true, true, true, true, true, true,
    CONFIG.COLORS.border,
    SpreadsheetApp.BorderStyle.SOLID
  );
  
  // Freeze header row
  sheet.setFrozenRows(1);
}

/**
 * Formats individual row with colors and styling
 */
function formatRow(sheet, row, rank, totalRows) {
  const rowRange = sheet.getRange(row, 1, 1, 4);
  
  // Determine background color
  let bgColor;
  if (rank === 1) {
    bgColor = CONFIG.COLORS.gold;
  } else if (rank === 2) {
    bgColor = CONFIG.COLORS.silver;
  } else if (rank === 3) {
    bgColor = CONFIG.COLORS.bronze;
  } else {
    // Alternate colors for remaining rows
    bgColor = (rank % 2 === 0) ? CONFIG.COLORS.evenRow : CONFIG.COLORS.oddRow;
  }
  
  rowRange.setBackground(bgColor);
  
  // Set row height
  sheet.setRowHeight(row, 35);
  
  // Font size based on rank
  let fontSize = 11;
  if (rank <= 3) {
    fontSize = 13;
    rowRange.setFontWeight('bold');
  }
  
  rowRange.setFontSize(fontSize);
  
  // Center align rank and numbers
  sheet.getRange(row, 1).setHorizontalAlignment('center'); // Rank
  sheet.getRange(row, 2).setHorizontalAlignment('left');   // Name
  sheet.getRange(row, 3).setHorizontalAlignment('center'); // Points
  sheet.getRange(row, 4).setHorizontalAlignment('center'); // Count
  
  // Vertical alignment
  rowRange.setVerticalAlignment('middle');
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Shows error message to user
 */
function showError(message) {
  SpreadsheetApp.getUi().alert(
    '❌ Error',
    message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Shows configuration dialog (for future enhancement)
 */
function showConfigDialog() {
  const html = `
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h2>⚙️ Leaderboard Configuration</h2>
      <p><strong>Current Settings:</strong></p>
      <ul>
        <li>Source Sheet: <code>${CONFIG.SOURCE_SHEET}</code></li>
        <li>Name Column: <code>${CONFIG.NAME_COLUMN}</code></li>
        <li>Total Column: <code>${CONFIG.TOTAL_COLUMN}</code></li>
        <li>Count Column: <code>${CONFIG.COUNT_COLUMN}</code></li>
        <li>Start Row: <code>${CONFIG.START_ROW}</code></li>
      </ul>
      <p>To change settings, edit the CONFIG object at the top of the script.</p>
      <p><em>Go to Extensions → Apps Script to edit.</em></p>
    </div>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(300);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Configuration');
}

// ============================================================================
// HTML LEADERBOARD FUNCTIONS
// ============================================================================

/**
 * Shows leaderboard as HTML in a modal dialog
 */
function showHtmlLeaderboard() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);
    
    if (!sourceSheet) {
      showError(`Source sheet "${CONFIG.SOURCE_SHEET}" not found!`);
      return;
    }
    
    const data = readTeamData(sourceSheet);
    if (data.length === 0) {
      showError('No data found in source sheet!');
      return;
    }
    
    const sortedData = sortTeams(data);
    const html = generateHtmlLeaderboard(sortedData);
    
    const htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(900)
      .setHeight(700);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🏆 Leaderboard');
    
  } catch (error) {
    showError('Error generating HTML leaderboard: ' + error.message);
  }
}

/**
 * Generates HTML page for leaderboard
 */
function generateHtmlLeaderboard(data) {
  let rowsHtml = '';
  
  for (let i = 0; i < data.length; i++) {
    const rank = i + 1;
    const team = data[i];
    const trophy = CONFIG.TROPHIES[rank] || '';
    const rankClass = rank <= 3 ? `rank-${rank}` : '';
    const animation = `animation-delay: ${i * 0.1}s;`;
    
    rowsHtml += `
      <tr class="team-row ${rankClass}" style="${animation}">
        <td class="rank-cell">
          <span class="rank-number">${rank}</span>
          ${trophy ? `<span class="trophy">${trophy}</span>` : ''}
        </td>
        <td class="team-name">${team.name}</td>
        <td class="points">
          <span class="points-value">${team.total}</span>
          <span class="points-label">bodů</span>
        </td>
        <td class="count">
          <span class="count-value">${team.count}</span>
          <span class="count-label">lidí</span>
        </td>
      </tr>
    `;
  }
  
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>🏆 Party Quiz Leaderboard</title>
  <style>
    * {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
    }
    
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      padding: 20px;
      min-height: 100vh;
    }
    
    .container {
      max-width: 1000px;
      margin: 0 auto;
    }
    
    .header {
      text-align: center;
      color: white;
      margin-bottom: 30px;
      animation: fadeInDown 0.8s ease;
    }
    
    .header h1 {
      font-size: 3em;
      margin-bottom: 10px;
      text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    }
    
    .header p {
      font-size: 1.2em;
      opacity: 0.9;
    }
    
    .leaderboard {
      background: white;
      border-radius: 15px;
      box-shadow: 0 10px 40px rgba(0,0,0,0.3);
      overflow: hidden;
      animation: fadeInUp 0.8s ease;
    }
    
    table {
      width: 100%;
      border-collapse: collapse;
    }
    
    thead {
      background: linear-gradient(135deg, #1a73e8 0%, #4285f4 100%);
      color: white;
    }
    
    thead th {
      padding: 20px;
      font-size: 1.1em;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 1px;
    }
    
    .team-row {
      border-bottom: 1px solid #f0f0f0;
      transition: all 0.3s ease;
      animation: slideIn 0.5s ease forwards;
      opacity: 0;
      transform: translateX(-20px);
    }
    
    .team-row:hover {
      background: #f8f9ff;
      transform: scale(1.02);
      box-shadow: 0 5px 15px rgba(0,0,0,0.1);
    }
    
    .team-row td {
      padding: 20px;
      font-size: 1.1em;
    }
    
    .rank-cell {
      text-align: center;
      font-weight: bold;
      font-size: 1.3em;
      width: 100px;
    }
    
    .trophy {
      font-size: 2em;
      display: block;
      margin-top: 5px;
      animation: bounce 1s infinite;
    }
    
    .rank-number {
      display: block;
      font-size: 0.9em;
      color: #666;
    }
    
    .team-name {
      font-weight: 600;
      font-size: 1.2em;
      color: #333;
    }
    
    .points, .count {
      text-align: center;
    }
    
    .points-value, .count-value {
      font-size: 1.5em;
      font-weight: bold;
      color: #1a73e8;
      display: block;
    }
    
    .points-label, .count-label {
      font-size: 0.8em;
      color: #999;
      text-transform: uppercase;
    }
    
    /* Medal colors for top 3 */
    .rank-1 {
      background: linear-gradient(135deg, #FFD700 0%, #FFA500 100%);
      color: #000;
    }
    
    .rank-1 .team-name {
      color: #000;
      font-size: 1.4em;
    }
    
    .rank-1 .points-value {
      color: #8B4513;
    }
    
    .rank-2 {
      background: linear-gradient(135deg, #C0C0C0 0%, #A8A8A8 100%);
      color: #000;
    }
    
    .rank-2 .team-name {
      color: #000;
      font-size: 1.3em;
    }
    
    .rank-2 .points-value {
      color: #505050;
    }
    
    .rank-3 {
      background: linear-gradient(135deg, #CD7F32 0%, #B87333 100%);
      color: #000;
    }
    
    .rank-3 .team-name {
      color: #000;
      font-size: 1.2em;
    }
    
    .rank-3 .points-value {
      color: #654321;
    }
    
    /* Animations */
    @keyframes fadeInDown {
      from {
        opacity: 0;
        transform: translateY(-30px);
      }
      to {
        opacity: 1;
        transform: translateY(0);
      }
    }
    
    @keyframes fadeInUp {
      from {
        opacity: 0;
        transform: translateY(30px);
      }
      to {
        opacity: 1;
        transform: translateY(0);
      }
    }
    
    @keyframes slideIn {
      to {
        opacity: 1;
        transform: translateX(0);
      }
    }
    
    @keyframes bounce {
      0%, 100% {
        transform: translateY(0);
      }
      50% {
        transform: translateY(-10px);
      }
    }
    
    /* Responsive */
    @media (max-width: 768px) {
      .header h1 {
        font-size: 2em;
      }
      
      .team-row td {
        padding: 15px 10px;
        font-size: 1em;
      }
      
      .team-name {
        font-size: 1em !important;
      }
    }
    
    /* Print styles */
    @media print {
      body {
        background: white;
      }
      
      .team-row:hover {
        transform: none;
      }
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>🏆 Leaderboard</h1>
      <p>Party Quiz Výsledky</p>
    </div>
    
    <div class="leaderboard">
      <table>
        <thead>
          <tr>
            <th>#</th>
            <th>Tým</th>
            <th>Body</th>
            <th>Počet lidí</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHtml}
        </tbody>
      </table>
    </div>
  </div>
</body>
</html>
  `;
}

/**
 * Serves the HTML leaderboard as a web app
 * This function is called when accessing the web app URL
 */
function doGet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);
    
    if (!sourceSheet) {
      return HtmlService.createHtmlOutput('<h1>Error: Source sheet not found</h1>');
    }
    
    const data = readTeamData(sourceSheet);
    if (data.length === 0) {
      return HtmlService.createHtmlOutput('<h1>Error: No data found</h1>');
    }
    
    const sortedData = sortTeams(data);
    const html = generateHtmlLeaderboard(sortedData);
    
    return HtmlService.createHtmlOutput(html)
      .setTitle('🏆 Party Quiz Leaderboard');
    
  } catch (error) {
    return HtmlService.createHtmlOutput('<h1>Error: ' + error.message + '</h1>');
  }
}

/**
 * Gets the web app URL for sharing
 */
function getWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  
  if (url) {
    const html = `
      <div style="font-family: Arial, sans-serif; padding: 20px;">
        <h2>🌐 Web App URL</h2>
        <p>Share this URL to display the leaderboard in any browser:</p>
        <div style="background: #f5f5f5; padding: 15px; margin: 15px 0; border-radius: 5px; word-break: break-all;">
          <strong>${url}</strong>
        </div>
        <p><a href="${url}" target="_blank">Open in new tab →</a></p>
        <hr style="margin: 20px 0;">
        <h3>How to deploy as web app:</h3>
        <ol>
          <li>Go to <strong>Extensions → Apps Script</strong></li>
          <li>Click <strong>Deploy → New deployment</strong></li>
          <li>Click gear icon → Select <strong>Web app</strong></li>
          <li>Set "Execute as" to <strong>Me</strong></li>
          <li>Set "Who has access" to <strong>Anyone</strong></li>
          <li>Click <strong>Deploy</strong></li>
          <li>Copy the Web app URL</li>
        </ol>
      </div>
    `;
    
    const htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(600)
      .setHeight(500);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Web App URL');
  } else {
    showError('Web app not deployed yet! Follow the instructions to deploy.');
  }
}

// ============================================================================
// EXAMPLE: Create sample data (for testing)
// ============================================================================

/**
 * Creates sample data for testing
 * Run this once to create test data
 */
function createSampleData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create or get source sheet
  let sheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SOURCE_SHEET);
  }
  
  // Sample data
  const data = [
    ['Tým', 'Body', 'Počet lidí'],
    ['Mozkovci', 85, 4],
    ['Rychlá Mysl', 92, 3],
    ['Geniální Parta', 78, 5],
    ['Chytré Hlavy', 92, 4],
    ['Kvízomani', 65, 3],
    ['Supermozek', 85, 2],
    ['Všeználkové', 55, 6]
  ];
  
  // Write data
  sheet.getRange(1, 1, data.length, 3).setValues(data);
  
  // Format header
  sheet.getRange(1, 1, 1, 3)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('white');
  
  SpreadsheetApp.getUi().alert('Sample data created in sheet: ' + CONFIG.SOURCE_SHEET);
}
