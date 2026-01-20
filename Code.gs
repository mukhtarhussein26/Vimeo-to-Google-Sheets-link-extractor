// Vimeo Video Source Link Extractor for Google Sheets
// Version with folder selection by name

// CONFIGURATION - EDIT THIS VALUE
const VIMEO_ACCESS_TOKEN = 'YOUR_VIMEO_ACCESS_TOKEN_HERE'; // Paste your token here

// Main function to extract video links from a specific folder
function extractFromSelectedFolder() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Get all folders/projects
    const folders = getAllFolders();
    
    if (folders.length === 0) {
      ui.alert('No folders found', 'You don\'t have any folders in your Vimeo account.', ui.ButtonSet.OK);
      return;
    }
    
    // Create a list of folder names for user to choose from
    const folderNames = folders.map((f, index) => `${index + 1}. ${f.name}`).join('\n');
    
    const response = ui.prompt(
      'Select a Folder',
      `Enter the number of the folder you want to extract videos from:\n\n${folderNames}\n\nEnter number (1-${folders.length}):`,
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const folderNumber = parseInt(response.getResponseText().trim());
    
    if (isNaN(folderNumber) || folderNumber < 1 || folderNumber > folders.length) {
      ui.alert('Invalid selection', 'Please enter a valid folder number.', ui.ButtonSet.OK);
      return;
    }
    
    const selectedFolder = folders[folderNumber - 1];
    
    // Confirm selection
    const confirmResponse = ui.alert(
      'Confirm Selection',
      `Extract videos from: "${selectedFolder.name}"?\n\nThis folder contains ${selectedFolder.metadata.connections.videos.total} videos.`,
      ui.ButtonSet.YES_NO
    );
    
    if (confirmResponse !== ui.Button.YES) {
      return;
    }
    
    // Extract videos from selected folder
    extractVideosFromFolder(selectedFolder);
    
  } catch (error) {
    ui.alert('Error: ' + error.message);
    Logger.log('Error details: ' + error);
  }
}

// Extract videos from all folders (alternative option)
function extractAllVideos() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'Extract All Videos',
    'This will extract ALL videos from your Vimeo account. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  extractVideos(null);
}

// Get all folders/projects from Vimeo
function getAllFolders() {
  const folders = [];
  let page = 1;
  let hasMore = true;
  
  while (hasMore) {
    const endpoint = `https://api.vimeo.com/me/projects?page=${page}&per_page=100`;
    
    const options = {
      method: 'get',
      headers: {
        'Authorization': `bearer ${VIMEO_ACCESS_TOKEN}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(endpoint, options);
    const statusCode = response.getResponseCode();
    
    if (statusCode !== 200) {
      throw new Error(`API Error (${statusCode}): ${response.getContentText()}`);
    }
    
    const result = JSON.parse(response.getContentText());
    folders.push(...result.data);
    
    hasMore = result.paging && result.paging.next !== null;
    page++;
  }
  
  return folders;
}

// Extract videos from a specific folder
function extractVideosFromFolder(folder) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  
  // Set up headers
  if (sheet.getRange(1, 1).getValue() !== 'Video Title') {
    sheet.getRange(1, 1, 1, 5).setValues([['Folder Name', 'Video Title', 'Video ID', 'Player URL', 'Upload Date']]);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  }
  
  // Clear existing data (except headers)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
  }
  
  try {
    const folderUri = folder.uri;
    const totalVideos = folder.metadata.connections.videos.total;
    
    if (totalVideos === 0) {
      ui.alert('No videos found in this folder.');
      return;
    }
    
    let currentRow = 2;
    let page = 1;
    let processedCount = 0;
    const batchSize = 50;
    
    while (processedCount < totalVideos) {
      const videos = getVideosFromFolder(folderUri, page, batchSize);
      
      if (videos.length === 0) break;
      
      const data = videos.map(video => [
        folder.name,
        video.name,
        video.uri.split('/').pop(),
        `https://player.vimeo.com/video/${video.uri.split('/').pop()}`,
        new Date(video.created_time)
      ]);
      
      sheet.getRange(currentRow, 1, data.length, 5).setValues(data);
      
      currentRow += data.length;
      processedCount += videos.length;
      page++;
      
      Logger.log(`Processed ${processedCount} of ${totalVideos} videos from "${folder.name}"`);
      Utilities.sleep(500);
    }
    
    // Format the sheet
    sheet.autoResizeColumns(1, 5);
    if (currentRow > 2) {
      sheet.getRange(2, 5, currentRow - 2, 1).setNumberFormat('yyyy-mm-dd hh:mm');
    }
    
    ui.alert(`Success!`, `Extracted ${processedCount} video links from "${folder.name}"`, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Error: ' + error.message);
    Logger.log('Error details: ' + error);
  }
}

// Get videos from a specific folder
function getVideosFromFolder(folderUri, page, perPage) {
  const endpoint = `https://api.vimeo.com${folderUri}/videos?page=${page}&per_page=${perPage}&sort=date&direction=asc`;
  
  const options = {
    method: 'get',
    headers: {
      'Authorization': `bearer ${VIMEO_ACCESS_TOKEN}`,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(endpoint, options);
  const statusCode = response.getResponseCode();
  
  // Handle "no more pages" gracefully (error 400 with code 2286)
  if (statusCode === 400) {
    const errorData = JSON.parse(response.getContentText());
    if (errorData.error_code === 2286) {
      // No more pages available - this is expected, return empty array
      return [];
    }
  }
  
  if (statusCode !== 200) {
    throw new Error(`API Error (${statusCode}): ${response.getContentText()}`);
  }
  
  const result = JSON.parse(response.getContentText());
  return result.data || [];
}

// Extract all videos (from entire account)
function extractVideos(folderUri) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  
  // Set up headers
  if (sheet.getRange(1, 1).getValue() !== 'Video Title') {
    sheet.getRange(1, 1, 1, 4).setValues([['Video Title', 'Video ID', 'Player URL', 'Upload Date']]);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  }
  
  // Clear existing data
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  }
  
  try {
    const totalVideos = getTotalVideoCount();
    
    if (totalVideos === 0) {
      ui.alert('No videos found.');
      return;
    }
    
    let currentRow = 2;
    let page = 1;
    let processedCount = 0;
    const batchSize = 50;
    
    while (processedCount < totalVideos) {
      const videos = getVideosBatch(page, batchSize);
      
      if (videos.length === 0) break;
      
      const data = videos.map(video => [
        video.name,
        video.uri.split('/').pop(),
        `https://player.vimeo.com/video/${video.uri.split('/').pop()}`,
        new Date(video.created_time)
      ]);
      
      sheet.getRange(currentRow, 1, data.length, 4).setValues(data);
      
      currentRow += data.length;
      processedCount += videos.length;
      page++;
      
      Logger.log(`Processed ${processedCount} of ${totalVideos} videos`);
      Utilities.sleep(500);
    }
    
    // Format the sheet
    sheet.autoResizeColumns(1, 4);
    if (currentRow > 2) {
      sheet.getRange(2, 4, currentRow - 2, 1).setNumberFormat('yyyy-mm-dd hh:mm');
    }
    
    ui.alert(`Successfully extracted ${processedCount} video links!`);
    
  } catch (error) {
    ui.alert('Error: ' + error.message);
    Logger.log('Error details: ' + error);
  }
}

// Get total video count
function getTotalVideoCount() {
  const endpoint = `https://api.vimeo.com/me/videos?per_page=1`;
  
  const options = {
    method: 'get',
    headers: {
      'Authorization': `bearer ${VIMEO_ACCESS_TOKEN}`,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(endpoint, options);
  const result = JSON.parse(response.getContentText());
  return result.total || 0;
}

// Get a batch of videos
function getVideosBatch(page, perPage) {
  const endpoint = `https://api.vimeo.com/me/videos?page=${page}&per_page=${perPage}`;
  
  const options = {
    method: 'get',
    headers: {
      'Authorization': `bearer ${VIMEO_ACCESS_TOKEN}`,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(endpoint, options);
  const result = JSON.parse(response.getContentText());
  return result.data || [];
}

// Create custom menu
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Vimeo Tools')
    .addItem('📁 Extract from Specific Folder', 'extractFromSelectedFolder')
    .addItem('📋 Extract All Videos', 'extractAllVideos')
    .addSeparator()
    .addItem('ℹ️ Help', 'showHelp')
    .addToUi();
}

// Help function
function showHelp() {
  const ui = SpreadshgetApp.getUi();
  const helpText = `Vimeo Video Link Extractor - Instructions:

SETUP:
1. Get API token from: https://developer.vimeo.com/apps
2. Paste it in line 5 of the script
3. Save and refresh the sheet

USAGE:
• "Extract from Specific Folder" - Shows a list of all your folders, select one
• "Extract All Videos" - Extracts all videos from your account

The script will create columns for:
- Folder name (when extracting from folder)
- Video title
- Video ID
- Player URL (the link you need!)
- Upload date

Player URL format: https://player.vimeo.com/video/[VIDEO_ID]`;
  
  ui.alert('Help', helpText, ui.ButtonSet.OK);
}
