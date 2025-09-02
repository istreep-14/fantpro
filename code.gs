// Function to create custom menu when spreadsheet opens
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Fantasy Tools')
    .addItem('Run All (Setup → Fetch → Master List)', 'runAllOperations')
    .addSeparator()
    .addItem('Setup Batch Scraper Config', 'setupBatchScraper')
    .addItem('Batch Process All URLs', 'batchProcessUrls')
    .addItem('Generate Master Rankings', 'runGenerateMasterRankings')
    .addSeparator()
    .addItem('Fetch Single RB Data (Example)', 'fetchFantasyProsData')
    .addToUi();
}

// Master function to run all operations at once
function runAllOperations() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Confirm action
  const response = ui.alert(
    'Run All Operations',
    'This will:\n' +
    '1. Setup/verify configuration\n' +
    '2. Fetch data from all configured URLs\n' +
    '3. Generate master rankings\n\n' +
    'This may take several minutes. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  const startTime = new Date();
  const results = {
    setup: false,
    fetched: 0,
    errors: [],
    masterGenerated: false
  };
  
  try {
    // Step 1: Setup configuration if not exists
    let configSheet;
    try {
      configSheet = spreadsheet.getSheetByName('Config');
      Logger.log('Config sheet already exists');
    } catch (e) {
      Logger.log('Creating config sheet...');
      setupBatchScraper();
      configSheet = spreadsheet.getSheetByName('Config');
      results.setup = true;
    }
    
    // Step 2: Run batch processing
    Logger.log('Starting batch processing...');
    const configData = configSheet.getDataRange().getValues();
    const totalUrls = configData.length - 1; // Subtract header row
    
    // Process each URL configuration
    for (let i = 1; i < configData.length; i++) {
      const tabName = configData[i][0];
      const url = configData[i][1];
      const headers = configData[i][2];
      const calculateTeam = configData[i][3] === 'true';
      
      if (!tabName || !url) continue;
      
      try {
        Logger.log(`Processing ${tabName}...`);
        const headerArray = headers.split(',').map(h => h.trim());
        const data = fetchAndParseData(url, headerArray, calculateTeam);
        
        if (data && data.length > 0) {
          createOrUpdateSheet(spreadsheet, tabName, data);
          results.fetched++;
          Logger.log(`Successfully processed ${tabName}: ${data.length} rows`);
        }
      } catch (error) {
        const errorMsg = `Failed to process ${tabName}: ${error.toString()}`;
        Logger.log(errorMsg);
        results.errors.push(errorMsg);
      }
      
      // Add a small delay to avoid hitting rate limits
      Utilities.sleep(1000);
    }
    
    // Step 3: Generate master rankings
    Logger.log('Generating master rankings...');
    generateMasterRankings(spreadsheet);
    results.masterGenerated = true;
    
    // Calculate execution time
    const endTime = new Date();
    const duration = Math.round((endTime - startTime) / 1000);
    
    // Show final results
    let message = 'Operations completed!\n\n';
    if (results.setup) {
      message += '✓ Configuration sheet created\n';
    }
    message += `✓ Fetched data from ${results.fetched} sources\n`;
    if (results.masterGenerated) {
      message += '✓ Master rankings generated\n';
    }
    message += `\nTotal time: ${duration} seconds\n`;
    
    if (results.errors.length > 0) {
      message += '\nErrors encountered:\n';
      results.errors.forEach(error => {
        message += `• ${error}\n`;
      });
    }
    
    message += '\nCheck the "Master Rankings" sheet for the final results.';
    
    ui.alert('Process Complete', message, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert(
      'Error',
      'A critical error occurred:\n' + error.toString(),
      ui.ButtonSet.OK
    );
    Logger.log('Critical error in runAllOperations: ' + error.toString());
  }
}

// Wrapper function for generateMasterRankings with UI feedback
function runGenerateMasterRankings() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  try {
    generateMasterRankings(spreadsheet);
    ui.alert(
      'Success!',
      'Master rankings have been generated successfully!\nCheck the "Master Rankings" sheet.',
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert(
      'Error',
      'Failed to generate master rankings: ' + error.toString(),
      ui.ButtonSet.OK
    );
  }
}

// Team mapping for defense names
const TEAM_MAPPINGS = {
  'Philadelphia Eagles': 'PHI',
  'Eagles': 'PHI',
  'Denver Broncos': 'DEN',
  'Broncos': 'DEN',
  'Buffalo Bills': 'BUF',
  'Bills': 'BUF',
  'Houston Texans': 'HOU',
  'Texans': 'HOU',
  'Baltimore Ravens': 'BAL',
  'Ravens': 'BAL',
  'Green Bay Packers': 'GB',
  'Packers': 'GB',
  'Pittsburgh Steelers': 'PIT',
  'Steelers': 'PIT',
  'Minnesota Vikings': 'MIN',
  'Vikings': 'MIN',
  'New York Giants': 'NYG',
  'Giants': 'NYG',
  'Detroit Lions': 'DET',
  'Lions': 'DET',
  'Los Angeles Rams': 'LAR',
  'Rams': 'LAR',
  'Los Angeles Chargers': 'LAC',
  'Chargers': 'LAC',
  'San Francisco 49ers': 'SF',
  '49ers': 'SF',
  'Washington Commanders': 'WAS',
  'Commanders': 'WAS',
  'Tampa Bay Buccaneers': 'TB',
  'Buccaneers': 'TB',
  'Dallas Cowboys': 'DAL',
  'Cowboys': 'DAL',
  'Kansas City Chiefs': 'KC',
  'Chiefs': 'KC',
  'Seattle Seahawks': 'SEA',
  'Seahawks': 'SEA',
  'Chicago Bears': 'CHI',
  'Bears': 'CHI',
  'Arizona Cardinals': 'ARI',
  'Cardinals': 'ARI',
  'Indianapolis Colts': 'IND',
  'Colts': 'IND',
  'Cleveland Browns': 'CLE',
  'Browns': 'CLE',
  'New York Jets': 'NYJ',
  'Jets': 'NYJ',
  'Cincinnati Bengals': 'CIN',
  'Bengals': 'CIN',
  'Las Vegas Raiders': 'LV',
  'Raiders': 'LV',
  'Miami Dolphins': 'MIA',
  'Dolphins': 'MIA',
  'Atlanta Falcons': 'ATL',
  'Falcons': 'ATL',
  'Jacksonville Jaguars': 'JAC',
  'Jaguars': 'JAC',
  'New Orleans Saints': 'NO',
  'Saints': 'NO',
  'Tennessee Titans': 'TEN',
  'Titans': 'TEN',
  'New England Patriots': 'NE',
  'Patriots': 'NE',
  'Carolina Panthers': 'CAR',
  'Panthers': 'CAR'
};

// Main function to set up the configuration sheet
function setupBatchScraper() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create or get the config sheet
  let configSheet;
  try {
    configSheet = spreadsheet.getSheetByName('Config');
    configSheet.clear(); // Clear existing content
  } catch (e) {
    configSheet = spreadsheet.insertSheet('Config');
  }
  
  // Set up the configuration table headers
  const configHeaders = ['Tab Name', 'URL', 'Headers (comma-separated)', 'Calculate Team'];
  configSheet.getRange(1, 1, 1, 4).setValues([configHeaders]);
  
  // Format the header row
  configSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  configSheet.getRange(1, 1, 1, 4).setBackground('#4285f4');
  configSheet.getRange(1, 1, 1, 4).setFontColor('#ffffff');
  
  // Add example data
  const exampleData = [
    ['RB_PPR', 'https://www.fantasypros.com/nfl/projections/rb.php?week=draft&scoring=PPR', 'Player,Team,ATT,YDS,TDS,REC,YDS,TDS,FL,FPTS', 'false'],
    ['WR_PPR', 'https://www.fantasypros.com/nfl/projections/wr.php?week=draft&scoring=PPR', 'Player,Team,REC,YDS,TDS,FL,FPTS', 'false'],
    ['QB_Standard', 'https://www.fantasypros.com/nfl/projections/qb.php?week=draft&scoring=STD', 'Player,Team,ATT,CMP,YDS,TDS,INT,FL,FPTS', 'false'],
    ['DEF_PPR', 'https://www.fantasypros.com/nfl/projections/dst.php?week=draft&scoring=PPR', 'Player,Team,FPTS', 'true']
  ];
  
  configSheet.getRange(2, 1, exampleData.length, 4).setValues(exampleData);
  
  // Auto-resize columns
  configSheet.autoResizeColumns(1, 4);
  
  // Set column widths for better visibility
  configSheet.setColumnWidth(1, 120); // Tab Name
  configSheet.setColumnWidth(2, 400); // URL
  configSheet.setColumnWidth(3, 300); // Headers
  configSheet.setColumnWidth(4, 100); // Calculate Team
  
  // Add instructions
  configSheet.getRange(exampleData.length + 3, 1, 1, 4).merge();
  configSheet.getRange(exampleData.length + 3, 1).setValue('Instructions: Fill in the rows above, then run "batchProcessUrls()" function. Set "Calculate Team" to "true" for defense sheets.');
  configSheet.getRange(exampleData.length + 3, 1).setBackground('#fff2cc');
  
  SpreadsheetApp.getUi().alert(
    'Setup Complete!', 
    'Configuration sheet created. Edit the URLs and headers above, then run the "batchProcessUrls" function.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// Main batch processing function
function batchProcessUrls() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = spreadsheet.getSheetByName('Config');
  
  if (!configSheet) {
    SpreadsheetApp.getUi().alert('Error', 'Config sheet not found. Run setupBatchScraper() first.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Get configuration data
  const configData = configSheet.getDataRange().getValues();
  const configs = [];
  
  // Parse config data (skip header row)
  for (let i = 1; i < configData.length; i++) {
    const row = configData[i];
    if (row[0] && row[1] && row[2]) { // Make sure all required fields are filled
      configs.push({
        tabName: row[0].toString().trim(),
        url: row[1].toString().trim(),
        headers: row[2].toString().split(',').map(h => h.trim()),
        calculateTeam: row[3] ? row[3].toString().toLowerCase() === 'true' : false
      });
    }
  }
  
  if (configs.length === 0) {
    SpreadsheetApp.getUi().alert('Error', 'No valid configurations found. Please fill in the Config sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Process each URL
  let successCount = 0;
  let errorCount = 0;
  const errors = [];
  
  for (const config of configs) {
    try {
      Logger.log(`Processing: ${config.tabName} - ${config.url}`);
      
      // Fetch and process data
      const data = fetchAndParseData(config.url, config.headers, config.calculateTeam);
      
      if (data.length > 1) { // More than just headers
        // Create or update sheet
        createOrUpdateSheet(spreadsheet, config.tabName, data);
        successCount++;
        Logger.log(`Success: ${config.tabName}`);
      } else {
        errors.push(`${config.tabName}: No data found`);
        errorCount++;
      }
      
      // Add small delay to avoid overwhelming the server
      Utilities.sleep(1000);
      
    } catch (error) {
      const errorMsg = `${config.tabName}: ${error.toString()}`;
      errors.push(errorMsg);
      errorCount++;
      Logger.log(`Error: ${errorMsg}`);
    }
  }
  
  // Show summary
  let message = `Batch processing complete!\nSuccess: ${successCount}\nErrors: ${errorCount}`;
  if (errors.length > 0) {
    message += '\n\nErrors:\n' + errors.join('\n');
  }
  
  // Generate master rankings
  try {
    generateMasterRankings(spreadsheet);
    message += '\n\nMaster Rankings sheet created successfully!';
  } catch (e) {
    message += '\n\nError creating Master Rankings: ' + e.toString();
  }
  
  SpreadsheetApp.getUi().alert('Batch Processing Results', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function fetchAndParseData(url, headers, calculateTeam = false) {
  // Fetch the HTML content
  const response = UrlFetchApp.fetch(url, {
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
  });
  
  const htmlContent = response.getContentText();
  
  // Parse the HTML to extract table data
  return parsePlayerData(htmlContent, headers, calculateTeam);
}

function parsePlayerData(htmlContent, customHeaders, calculateTeam = false) {
  const data = [customHeaders];
  
  // Look for the main data table - Fantasy Pros uses specific table structure
  const tablePattern = /<table[^>]*class="[^"]*table[^"]*"[^>]*>.*?<\/table>/gis;
  const tables = htmlContent.match(tablePattern) || [];
  
  let foundData = false;
  
  for (let table of tables) {
    const rowPattern = /<tr[^>]*>.*?<\/tr>/gis;
    const rows = table.match(rowPattern) || [];
    
    for (let i = 1; i < rows.length; i++) { // Skip header row
      const row = rows[i];
      const cellPattern = /<td[^>]*>(.*?)<\/td>/gis;
      const cells = [];
      let match;
      
      // Reset regex
      cellPattern.lastIndex = 0;
      
      while ((match = cellPattern.exec(row)) !== null) {
        let cellContent = match[1];
        // Clean HTML tags
        cellContent = cellContent.replace(/<[^>]*>/g, '').trim();
        // Remove line breaks and extra whitespace
        cellContent = cellContent.replace(/\s+/g, ' ').trim();
        cells.push(cellContent);
      }
      
      if (cells.length >= 2) {
        const processedRow = processPlayerRow(cells, customHeaders.length, calculateTeam);
        if (processedRow && processedRow.length >= Math.min(customHeaders.length, 3)) {
          // Ensure exact number of columns as headers
          while (processedRow.length < customHeaders.length) {
            processedRow.push('');
          }
          data.push(processedRow.slice(0, customHeaders.length));
          foundData = true;
        }
      }
    }
    
    if (foundData) break;
  }
  
  // If no data found, try alternative parsing
  if (!foundData) {
    Logger.log('Primary parsing failed, trying alternative method...');
    tryAlternativeParsing(htmlContent, data, customHeaders, calculateTeam);
  }
  
  return data;
}

function processPlayerRow(cells, expectedColumns, calculateTeam = false) {
  if (cells.length === 0) return null;
  
  // The first cell typically contains player name and team concatenated
  const firstCell = cells[0];
  
  // Separate player name from team using regex
  // Team abbreviations are typically 2-4 capital letters at the end
  const playerTeamMatch = firstCell.match(/^(.+?)([A-Z]{2,4})$/);
  
  let playerName, team;
  
  if (playerTeamMatch) {
    playerName = playerTeamMatch[1].trim();
    team = playerTeamMatch[2];
    
    // Remove any ranking numbers at the beginning of player name
    playerName = playerName.replace(/^\d+\.\s*/, '');
  } else {
    // Fallback: try to split by common patterns
    const words = firstCell.split(/\s+/);
    const lastWord = words[words.length - 1];
    
    if (lastWord && lastWord.length >= 2 && lastWord === lastWord.toUpperCase()) {
      team = lastWord;
      playerName = words.slice(0, -1).join(' ');
      // Remove ranking numbers
      playerName = playerName.replace(/^\d+\.\s*/, '');
    } else {
      // If we can't separate, put everything in player name
      playerName = firstCell.replace(/^\d+\.\s*/, '');
      team = '';
    }
  }
  
  // If calculateTeam is true and team is empty, try to find team from name
  if (calculateTeam && !team) {
    team = findTeamFromName(playerName);
  }
  
  // Build the processed row
  const processedRow = [playerName, team];
  
  // Add the remaining stats (skip the first cell since we processed it)
  for (let i = 1; i < cells.length && processedRow.length < expectedColumns; i++) {
    let stat = cells[i];
    
    // Clean up stat values
    stat = stat.replace(/[^\d.-]/g, ''); // Keep only numbers, decimals, and negative signs
    if (stat === '' || isNaN(stat)) {
      stat = '0';
    }
    
    processedRow.push(stat);
  }
  
  return processedRow;
}

// Function to find team abbreviation from player/defense name
function findTeamFromName(playerName) {
  // Check if the name matches any team in our mapping
  for (const [teamName, abbreviation] of Object.entries(TEAM_MAPPINGS)) {
    if (playerName.toLowerCase().includes(teamName.toLowerCase())) {
      return abbreviation;
    }
  }
  return '';
}

function tryAlternativeParsing(htmlContent, data, customHeaders, calculateTeam = false) {
  // Look for any table-like structure or try to find player data patterns
  const numColumns = customHeaders.length - 2; // Subtract Player and Team columns
  const statPattern = '\\s*([\\d.-]+)'.repeat(numColumns);
  const playerPattern = new RegExp(`([A-Z][a-z]+ [A-Z][a-z]+(?:\\s[A-Z][a-z]+)*)\\s*([A-Z]{2,4})${statPattern}`, 'g');
  
  let match;
  while ((match = playerPattern.exec(htmlContent)) !== null) {
    const row = [match[1], match[2]]; // Player name and team
    
    // Add remaining stats
    for (let i = 3; i < match.length; i++) {
      row.push(match[i] || '0');
    }
    
    // Ensure correct number of columns
    while (row.length < customHeaders.length) {
      row.push('0');
    }
    
    data.push(row.slice(0, customHeaders.length));
  }
}

function createOrUpdateSheet(spreadsheet, tabName, data) {
  // Try to get existing sheet or create new one
  let sheet;
  try {
    sheet = spreadsheet.getSheetByName(tabName);
    sheet.clear(); // Clear existing data
  } catch (e) {
    sheet = spreadsheet.insertSheet(tabName);
  }
  
  // Add data to sheet
  if (data.length > 0) {
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    
    // Format header row
    sheet.getRange(1, 1, 1, data[0].length).setFontWeight('bold');
    sheet.getRange(1, 1, 1, data[0].length).setBackground('#e6f3ff');
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, data[0].length);
    
    // Add timestamp
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    sheet.getRange(data.length + 2, 1).setValue(`Last updated: ${timestamp}`);
    sheet.getRange(data.length + 2, 1).setFontStyle('italic');
    sheet.getRange(data.length + 2, 1).setFontColor('#666666');
  }
  
  return sheet;
}

// Convenience function for single URL processing (backward compatibility)
function fetchFantasyProsData() {
  const url = 'https://www.fantasypros.com/nfl/projections/rb.php?week=draft&scoring=PPR&week=draft';
  const headers = ['Player', 'Team', 'ATT', 'YDS', 'TDS', 'REC', 'YDS', 'TDS', 'FL', 'FPTS'];
  
  try {
    const data = fetchAndParseData(url, headers);
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    createOrUpdateSheet(spreadsheet, 'RB_Data', data);
    
    SpreadsheetApp.getUi().alert(
      'Success!', 
      'Fantasy Pros RB data has been imported successfully!',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('Error fetching data: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', 'Failed to fetch data: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// Function to generate master rankings from all player sheets
function generateMasterRankings(spreadsheet) {
  const allPlayers = [];
  const sheets = spreadsheet.getSheets();
  const configSheet = spreadsheet.getSheetByName('Config');
  
  // Get position names from config sheet
  const configData = configSheet.getDataRange().getValues();
  const positionMap = {};
  
  for (let i = 1; i < configData.length; i++) {
    const tabName = configData[i][0];
    if (tabName) {
      // Extract position from tab name (e.g., RB_PPR -> RB)
      const position = tabName.split('_')[0];
      positionMap[tabName] = position;
    }
  }
  
  // Collect all players from each sheet
  for (const sheet of sheets) {
    const sheetName = sheet.getName();
    
    // Skip non-player sheets
    if (sheetName === 'Config' || sheetName === 'Master Rankings') {
      continue;
    }
    
    const position = positionMap[sheetName] || sheetName;
    const data = sheet.getDataRange().getValues();
    
    // Find FPTS column index
    const headers = data[0];
    const fptsIndex = headers.indexOf('FPTS');
    
    if (fptsIndex === -1) {
      Logger.log(`No FPTS column found in sheet: ${sheetName}`);
      continue;
    }
    
    // Process each player row (skip header)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const playerName = row[0];
      const team = row[1] || '';
      const fpts = parseFloat(row[fptsIndex]) || 0;
      
      // Skip empty rows or rows without valid player names
      if (!playerName || playerName.toString().trim() === '' || 
          playerName.toString().includes('Last updated')) {
        continue;
      }
      
      allPlayers.push({
        name: playerName.toString().trim(),
        team: team.toString().trim(),
        position: position,
        fpts: fpts,
        // Add a unique identifier for tie-breaking (sheet name + row index)
        tieBreaker: `${sheetName}_${i.toString().padStart(4, '0')}`
      });
    }
  }
  
  // Sort players by FPTS (descending) with tie-breaking
  allPlayers.sort((a, b) => {
    if (b.fpts !== a.fpts) {
      return b.fpts - a.fpts;
    }
    // Tie-breaker: use position priority (QB > RB > WR > TE > DEF > K)
    const positionPriority = {
      'QB': 1,
      'RB': 2,
      'WR': 3,
      'TE': 4,
      'DEF': 5,
      'K': 6
    };
    
    const aPriority = positionPriority[a.position] || 99;
    const bPriority = positionPriority[b.position] || 99;
    
    if (aPriority !== bPriority) {
      return aPriority - bPriority;
    }
    
    // Final tie-breaker: alphabetical by name
    return a.name.localeCompare(b.name);
  });
  
  // Create Master Rankings sheet
  let masterSheet;
  try {
    masterSheet = spreadsheet.getSheetByName('Master Rankings');
    masterSheet.clear();
  } catch (e) {
    masterSheet = spreadsheet.insertSheet('Master Rankings');
  }
  
  // Set headers
  const headers = ['Rank', 'Name', 'Team', 'Pos', 'FPTS'];
  masterSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format header row
  masterSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  masterSheet.getRange(1, 1, 1, headers.length).setBackground('#4285f4');
  masterSheet.getRange(1, 1, 1, headers.length).setFontColor('#ffffff');
  
  // Add player data
  const playerData = allPlayers.map((player, index) => [
    index + 1, // Rank
    player.name,
    player.team,
    player.position,
    player.fpts
  ]);
  
  if (playerData.length > 0) {
    masterSheet.getRange(2, 1, playerData.length, headers.length).setValues(playerData);
    
    // Format FPTS column as numbers with 2 decimal places
    masterSheet.getRange(2, 5, playerData.length, 1).setNumberFormat('#,##0.00');
    
    // Auto-resize columns
    masterSheet.autoResizeColumns(1, headers.length);
    
    // Add timestamp
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    masterSheet.getRange(playerData.length + 3, 1).setValue(`Last updated: ${timestamp}`);
    masterSheet.getRange(playerData.length + 3, 1).setFontStyle('italic');
    masterSheet.getRange(playerData.length + 3, 1).setFontColor('#666666');
    
    // Add alternating row colors for better readability
    const range = masterSheet.getRange(2, 1, playerData.length, headers.length);
    range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  }
  
  Logger.log(`Master Rankings created with ${playerData.length} players`);
}
