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
    .addItem('🌐 View as HTML Page', 'showHtmlLeaderboard')
    .addItem('📱 Get HTML Page URL', 'getWebAppUrl')
    .addToUi();
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
      .setWidth(2000)
      .setHeight(1200);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🏆 Leaderboard');
    
  } catch (error) {
    showError('Error generating HTML leaderboard: ' + error.message);
  }
}

/**
 * Groups teams by rank (handling ties)
 */
function groupTeamsByRank(data) {
  const groups = [];
  let currentRank = 1;
  let i = 0;
  
  while (i < data.length) {
    const group = {
      rank: currentRank,
      teams: [data[i]]
    };
    
    // Check for ties (same total and count)
    let j = i + 1;
    while (j < data.length && 
           data[j].total === data[i].total && 
           data[j].count === data[i].count) {
      group.teams.push(data[j]);
      j++;
    }
    
    groups.push(group);
    currentRank += group.teams.length;
    i = j;
  }
  
  return groups;
}

/**
 * Generates HTML page for leaderboard
 */
function generateHtmlLeaderboard(data) {
  // Prepare data as JSON for JavaScript
  const groups = groupTeamsByRank(data);
  const dataJson = JSON.stringify(groups);
  
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
      min-height: 100vh;
      display: flex;
      flex-direction: column;
      overflow: hidden;
    }
    
    .container {
      flex: 1;
      display: flex;
      flex-direction: column;
      padding: 10px;
      max-height: 100vh;
    }
    
    .header {
      text-align: center;
      color: white;
      margin-bottom: 10px;
      animation: fadeInDown 0.8s ease;
    }
    
    .header h1 {
      font-size: 2.5em;
      margin-bottom: 5px;
      text-shadow: 3px 3px 6px rgba(0,0,0,0.4);
    }
    
    .header p {
      font-size: 1.1em;
      opacity: 0.9;
    }
    
    /* Mode Controls */
    .controls {
      text-align: center;
      margin-bottom: 10px;
      animation: fadeIn 0.8s ease;
    }
    
    .mode-btn {
      background: white;
      border: none;
      padding: 10px 20px;
      margin: 0 10px;
      border-radius: 20px;
      font-size: 0.95em;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.3s ease;
      box-shadow: 0 4px 15px rgba(0,0,0,0.2);
    }
    
    .mode-btn:hover {
      transform: translateY(-2px);
      box-shadow: 0 6px 20px rgba(0,0,0,0.3);
    }
    
    .mode-btn.active {
      background: #1a73e8;
      color: white;
    }
    
    /* Navigation hint for reveal mode */
    .nav-hint {
      text-align: center;
      color: white;
      font-size: 1em;
      margin-bottom: 10px;
      animation: pulse 2s infinite;
      display: none;
    }
    
    .nav-hint.show {
      display: block;
    }
    
    /* Leaderboard Container */
    .leaderboard {
      flex: 1;
      background: white;
      border-radius: 20px;
      box-shadow: 0 15px 50px rgba(0,0,0,0.4);
      overflow: auto;
      animation: fadeInUp 0.8s ease;
      padding: 20px;
    }
    
    /* Two column grid for show all mode */
    .mode-all #leaderboard-body {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 20px;
      padding: 0;
    }
    
    .leaderboard-column {
      display: flex;
      flex-direction: column;
      gap: 15px;
    }
    
    /* Single column for reveal mode */
    .mode-reveal #leaderboard-body {
      display: flex;
      flex-direction: column;
    }
    
    .mode-reveal .leaderboard-column {
      width: 100%;
    }
    
    .mode-reveal #column-2 {
      display: none;
    }
    
    /* Table styles */
    table {
      width: 100%;
      border-collapse: collapse;
    }
    
    thead {
      background: linear-gradient(135deg, #1a73e8 0%, #4285f4 100%);
      color: white;
      display: none; /* Hide in grid mode, show headers visually in cards */
    }
    
    .mode-reveal thead {
      display: table-header-group; /* Show in reveal mode */
    }
    
    thead th {
      padding: 12px;
      font-size: 1em;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 1px;
    }
    
    tbody {
      overflow-y: auto;
      flex: 1;
    }
    
    /* Team row styles */
    .team-row {
      border: 2px solid #e0e0e0;
      border-radius: 10px;
      transition: all 0.3s ease;
      opacity: 0;
      padding: 8px 12px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      background: white;
      min-height: 45px;
    }
    
    /* Show all mode - rows visible immediately */
    .mode-all .team-row {
      animation: slideIn 0.5s ease forwards;
    }
    
    /* Reveal mode - rows hidden by default, display as full width */
    .mode-reveal .team-row {
      opacity: 0;
      display: none;
      width: 100%;
    }
    
    .mode-reveal .team-row.revealed {
      display: flex;
      animation: revealRow 0.8s ease forwards;
    }
    
    .team-row:hover {
      background: #f8f9ff;
      transform: scale(1.02);
      box-shadow: 0 3px 10px rgba(0,0,0,0.15);
    }
    
    /* Team row content layout */
    .team-row-content {
      display: flex;
      align-items: center;
      width: 100%;
      gap: 10px;
    }
    
    .rank-cell {
      text-align: center;
      font-weight: bold;
      font-size: 1em;
      min-width: 50px;
      display: flex;
      flex-direction: column;
      align-items: center;
      flex-shrink: 0;
    }
    
    .trophy {
      font-size: 1.5em;
      line-height: 1;
    }
    
    .rank-number {
      font-size: 0.85em;
      color: #666;
    }
    
    .team-name {
      font-weight: 600;
      font-size: 0.95em;
      color: #333;
      flex: 1;
      padding: 0 8px;
    }
    
    .points, .count {
      text-align: center;
      min-width: 60px;
      flex-shrink: 0;
    }
    
    .points-value, .count-value {
      font-size: 1.1em;
      font-weight: bold;
      color: #1a73e8;
      display: block;
      line-height: 1.2;
    }
    
    .points-label, .count-label {
      font-size: 0.65em;
      color: #999;
      text-transform: uppercase;
      line-height: 1;
    }
    
    /* Medal colors for top 3 */
    .rank-1 {
      background: linear-gradient(135deg, #FFD700 0%, #FFA500 100%);
      color: #000;
      border: 3px solid #DAA520;
    }
    
    .rank-1 .team-name {
      color: #000;
      font-size: 1.05em;
      font-weight: 800;
    }
    
    .rank-1 .points-value {
      color: #8B4513;
      font-size: 1.3em;
    }
    
    .rank-2 {
      background: linear-gradient(135deg, #C0C0C0 0%, #A8A8A8 100%);
      color: #000;
      border: 3px solid #909090;
    }
    
    .rank-2 .team-name {
      color: #000;
      font-size: 1.05em;
      font-weight: 700;
    }
    
    .rank-2 .points-value {
      color: #505050;
      font-size: 1.25em;
    }
    
    .rank-3 {
      background: linear-gradient(135deg, #CD7F32 0%, #B87333 100%);
      color: #000;
      border: 3px solid #A0522D;
    }
    
    .rank-3 .team-name {
      color: #000;
      font-size: 1.05em;
      font-weight: 700;
    }
    
    .rank-3 .points-value {
      color: #654321;
      font-size: 1.2em;
    }
    
    /* Tied teams highlight */
    .tied-group {
      background: #fff9e6;
      border: 2px solid #ffdb58;
    }
    
    /* Animations */
    @keyframes fadeInDown {
      from {
        opacity: 0;
        transform: translateY(-50px);
      }
      to {
        opacity: 1;
        transform: translateY(0);
      }
    }
    
    @keyframes fadeInUp {
      from {
        opacity: 0;
        transform: translateY(50px);
      }
      to {
        opacity: 1;
        transform: translateY(0);
      }
    }
    
    @keyframes fadeIn {
      from {
        opacity: 0;
      }
      to {
        opacity: 1;
      }
    }
    
    @keyframes slideIn {
      to {
        opacity: 1;
      }
    }
    
    @keyframes revealRow {
      0% {
        opacity: 0;
        transform: translateY(30px) scale(0.9);
      }
      60% {
        transform: translateY(-10px) scale(1.05);
      }
      100% {
        opacity: 1;
        transform: translateY(0) scale(1);
      }
    }
    
    @keyframes bounce {
      0%, 100% {
        transform: translateY(0);
      }
      50% {
        transform: translateY(-15px);
      }
    }
    
    @keyframes pulse {
      0%, 100% {
        opacity: 1;
      }
      50% {
        opacity: 0.5;
      }
    }
    
    /* Responsive */
    @media (max-width: 768px) {
      .header h1 {
        font-size: 2.5em;
      }
      
      .team-row td {
        padding: 15px 10px;
        font-size: 1.1em;
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
    
    <div class="controls">
      <button class="mode-btn active" onclick="setMode('all')">📊 Zobrazit vše</button>
      <button class="mode-btn" onclick="setMode('reveal')">🎬 Finálové odhalení</button>
    </div>
    
    <div class="nav-hint">
      ⬆️ Šipka nahoru: Další tým/y ⬆️
    </div>
    
    <div class="leaderboard mode-all">
      <div id="leaderboard-body">
      </div>
    </div>
  </div>
  
  <script>
    // Data from server
    const groups = ${dataJson};
    const trophies = { 1: '🥇', 2: '🥈', 3: '🥉' };
    
    let currentMode = 'all';
    let revealedGroupIndex = -1;
    
    // Initialize
    document.addEventListener('DOMContentLoaded', function() {
      renderLeaderboard();
      setupKeyboardNavigation();
    });
    
    function renderLeaderboard() {
      const tbody = document.getElementById('leaderboard-body');
      tbody.innerHTML = '';
      
      // In reveal mode, use single column. In show-all mode, use two columns
      if (currentMode === 'reveal') {
        // Single column for reveal mode - all teams in one container
        let rowIndex = 0;
        groups.forEach((group, groupIndex) => {
          const isTied = group.teams.length > 1;
          
          group.teams.forEach((team, teamIndex) => {
            const div = document.createElement('div');
            div.className = 'team-row';
            div.dataset.groupIndex = groupIndex;
            div.dataset.rowIndex = rowIndex;
            
            // Add rank class for top 3
            if (group.rank <= 3 && !isTied) {
              div.classList.add(\`rank-\${group.rank}\`);
            }
            
            // Add tied group class
            if (isTied) {
              div.classList.add('tied-group');
            }
            
            const trophy = trophies[group.rank] || '';
            
            div.innerHTML = \`
              <div class="team-row-content">
                <div class="rank-cell">
                  \${trophy ? \`<div class="trophy">\${trophy}</div>\` : ''}
                  <div class="rank-number">\${group.rank}</div>
                </div>
                <div class="team-name">\${team.name}</div>
                <div class="points">
                  <div class="points-value">\${team.total}</div>
                  <div class="points-label">bodů</div>
                </div>
                <div class="count">
                  <div class="count-value">\${team.count}</div>
                  <div class="count-label">lidí</div>
                </div>
              </div>
            \`;
            
            tbody.appendChild(div);
            rowIndex++;
          });
        });
      } else {
        // Two columns for show-all mode
        const column1 = document.createElement('div');
        column1.className = 'leaderboard-column';
        column1.id = 'column-1';
        
        const column2 = document.createElement('div');
        column2.className = 'leaderboard-column';
        column2.id = 'column-2';
        
        tbody.appendChild(column1);
        tbody.appendChild(column2);
        
        // Calculate total number of teams for even split
        let totalTeams = 0;
        groups.forEach(group => {
          totalTeams += group.teams.length;
        });
        const halfPoint = Math.ceil(totalTeams / 2);
        
        let rowIndex = 0;
        groups.forEach((group, groupIndex) => {
          const isTied = group.teams.length > 1;
          
          group.teams.forEach((team, teamIndex) => {
            const div = document.createElement('div');
            div.className = 'team-row';
            div.dataset.groupIndex = groupIndex;
            div.dataset.rowIndex = rowIndex;
            
            // Add rank class for top 3
            if (group.rank <= 3 && !isTied) {
              div.classList.add(\`rank-\${group.rank}\`);
            }
            
            // Add tied group class
            if (isTied) {
              div.classList.add('tied-group');
            }
            
            // Animation delay for "show all" mode
            div.style.animationDelay = \`\${rowIndex * 0.05}s\`;
            
            const trophy = trophies[group.rank] || '';
            
            div.innerHTML = \`
              <div class="team-row-content">
                <div class="rank-cell">
                  \${trophy ? \`<div class="trophy">\${trophy}</div>\` : ''}
                  <div class="rank-number">\${group.rank}</div>
                </div>
                <div class="team-name">\${team.name}</div>
                <div class="points">
                  <div class="points-value">\${team.total}</div>
                  <div class="points-label">bodů</div>
                </div>
                <div class="count">
                  <div class="count-value">\${team.count}</div>
                  <div class="count-label">lidí</div>
                </div>
              </div>
            \`;
            
            // Split teams evenly: first half in column 1, second half in column 2
            if (rowIndex < halfPoint) {
              column1.appendChild(div);
            } else {
              column2.appendChild(div);
            }
            
            rowIndex++;
          });
        });
      }
    }
    
    function setMode(mode) {
      currentMode = mode;
      const leaderboard = document.querySelector('.leaderboard');
      const navHint = document.querySelector('.nav-hint');
      const buttons = document.querySelectorAll('.mode-btn');
      
      // Update button states
      buttons.forEach(btn => btn.classList.remove('active'));
      event.target.classList.add('active');
      
      // Update leaderboard class
      leaderboard.className = 'leaderboard mode-' + mode;
      
      if (mode === 'all') {
        navHint.classList.remove('show');
        // Re-render to show in columns
        renderLeaderboard();
      } else if (mode === 'reveal') {
        navHint.classList.add('show');
        // Re-render for single column reveal mode
        renderLeaderboard();
        // Hide all rows, start from bottom
        revealedGroupIndex = -1;
        document.querySelectorAll('.team-row').forEach(row => {
          row.classList.remove('revealed');
        });
      }
    }
    
    function setupKeyboardNavigation() {
      document.addEventListener('keydown', function(e) {
        if (currentMode !== 'reveal') return;
        
        // Up arrow or Space - reveal next group
        if (e.key === 'ArrowUp' || e.key === ' ') {
          e.preventDefault();
          revealNextGroup();
        }
        
        // Down arrow - go back one group
        if (e.key === 'ArrowDown') {
          e.preventDefault();
          hideLastGroup();
        }
        
        // R key - reset
        if (e.key === 'r' || e.key === 'R') {
          e.preventDefault();
          resetReveal();
        }
      });
    }
    
    function revealNextGroup() {
      if (revealedGroupIndex >= groups.length - 1) {
        return; // All revealed
      }
      
      revealedGroupIndex++;
      
      // Reveal from bottom to top, so we start from the end
      const groupToReveal = groups.length - 1 - revealedGroupIndex;
      
      // Find all rows in this group and reveal them
      document.querySelectorAll(\`[data-group-index="\${groupToReveal}"]\`).forEach(row => {
        row.classList.add('revealed');
      });
      
      // Hide hint when all revealed
      if (revealedGroupIndex >= groups.length - 1) {
        document.querySelector('.nav-hint').textContent = '🎉 Všichni odhaleni! 🎉';
      }
    }
    
    function hideLastGroup() {
      if (revealedGroupIndex < 0) return;
      
      const groupToHide = groups.length - 1 - revealedGroupIndex;
      
      document.querySelectorAll(\`[data-group-index="\${groupToHide}"]\`).forEach(row => {
        row.classList.remove('revealed');
      });
      
      revealedGroupIndex--;
      
      // Restore hint
      document.querySelector('.nav-hint').textContent = '⬆️ Šipka nahoru: Další tým/y ⬆️';
    }
    
    function resetReveal() {
      revealedGroupIndex = -1;
      document.querySelectorAll('.team-row').forEach(row => {
        row.classList.remove('revealed');
      });
      document.querySelector('.nav-hint').textContent = '⬆️ Šipka nahoru: Další tým/y ⬆️';
    }
  </script>
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
