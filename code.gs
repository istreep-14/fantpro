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
  const configHeaders = ['Tab Name', 'URL', 'Headers (comma-separated)'];
  configSheet.getRange(1, 1, 1, 3).setValues([configHeaders]);
  
  // Format the header row
  configSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  configSheet.getRange(1, 1, 1, 3).setBackground('#4285f4');
  configSheet.getRange(1, 1, 1, 3).setFontColor('#ffffff');
  
  // Add example data
  const exampleData = [
    ['RB_PPR', 'https://www.fantasypros.com/nfl/projections/rb.php?week=draft&scoring=PPR', 'Player,Team,ATT,YDS,TDS,REC,YDS,TDS,FL,FPTS'],
    ['WR_PPR', 'https://www.fantasypros.com/nfl/projections/wr.php?week=draft&scoring=PPR', 'Player,Team,REC,YDS,TDS,FL,FPTS'],
    ['QB_Standard', 'https://www.fantasypros.com/nfl/projections/qb.php?week=draft&scoring=STD', 'Player,Team,ATT,CMP,YDS,TDS,INT,FL,FPTS']
  ];
  
  configSheet.getRange(2, 1, exampleData.length, 3).setValues(exampleData);
  
  // Auto-resize columns
  configSheet.autoResizeColumns(1, 3);
  
  // Set column widths for better visibility
  configSheet.setColumnWidth(1, 120); // Tab Name
  configSheet.setColumnWidth(2, 400); // URL
  configSheet.setColumnWidth(3, 300); // Headers
  
  // Add instructions
  configSheet.getRange(exampleData.length + 3, 1, 1, 3).merge();
  configSheet.getRange(exampleData.length + 3, 1).setValue('Instructions: Fill in the rows above, then run "batchProcessUrls()" function');
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
        headers: row[2].toString().split(',').map(h => h.trim())
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
      const data = fetchAndParseData(config.url, config.headers);
      
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
  
  SpreadsheetApp.getUi().alert('Batch Processing Results', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function fetchAndParseData(url, headers) {
  // Fetch the HTML content
  const response = UrlFetchApp.fetch(url, {
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
  });
  
  const htmlContent = response.getContentText();
  
  // Parse the HTML to extract table data
  return parsePlayerData(htmlContent, headers);
}

function parsePlayerData(htmlContent, customHeaders) {
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
        const processedRow = processPlayerRow(cells, customHeaders.length);
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
    tryAlternativeParsing(htmlContent, data, customHeaders);
  }
  
  return data;
}

function processPlayerRow(cells, expectedColumns) {
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

function tryAlternativeParsing(htmlContent, data, customHeaders) {
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
