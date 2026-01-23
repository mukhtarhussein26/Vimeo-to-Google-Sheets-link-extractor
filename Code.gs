// Vimeo Video Link Extractor & Thumbnail Uploader for Google Sheets
// Automation tool for Vimeo video management

// CONFIGURATION - EDIT THESE VALUES
const VIMEO_ACCESS_TOKEN = 'YOUR_VIMEO_ACCESS_TOKEN_HERE'; // Paste your token here (needs Public, Private, Edit, Upload scopes)
const DRIVE_FOLDER_ID = 'YOUR_GOOGLE_DRIVE_FOLDER_ID_HERE'; // Folder with thumbnail images

// ==================== VIDEO EXTRACTION FUNCTIONS ====================

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
    
    // Get actual video count by fetching first page
    const actualCount = getActualVideoCount(selectedFolder.uri);
    
    // Confirm selection
    const confirmResponse = ui.alert(
      'Confirm Selection',
      `Extract videos from: "${selectedFolder.name}"?\n\nThis folder contains ${actualCount} videos.`,
      ui.ButtonSet.YES_NO
    );
    
    if (confirmResponse !== ui.Button.YES) {
      return;
    }
    
    // Extract videos from selected folder
    extractVideosFromFolder(selectedFolder, actualCount);
    
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

// Get actual video count from folder
function getActualVideoCount(folderUri) {
  const endpoint = `https://api.vimeo.com${folderUri}/videos?per_page=1`;
  
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

// Extract videos from a specific folder
function extractVideosFromFolder(folder, totalVideos) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  
  // Set up headers
  if (sheet.getRange(1, 1).getValue() !== 'Folder Name') {
    sheet.getRange(1, 1, 1, 5).setValues([['Folder Name', 'Video Title', 'Video ID', 'Player URL', 'Upload Date']]);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  }
  
  // Clear existing data (except headers)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  
  try {
    const folderUri = folder.uri;
    
    if (totalVideos === 0) {
      ui.alert('No videos found in this folder.');
      return;
    }
    
    let currentRow = 2;
    let page = 1;
    let processedCount = 0;
    const batchSize = 100; // Increased batch size
    
    // Keep fetching until we get all videos or hit empty response
    while (true) {
      const videos = getVideosFromFolder(folderUri, page, batchSize);
      
      if (videos.length === 0) {
        Logger.log(`No more videos found at page ${page}`);
        break;
      }
      
      const data = videos.map(video => [
        folder.name,
        video.name,
        video.uri.split('/').pop(),
        `https://player.vimeo.com/video/${video.uri.split('/').pop()}?`,
        new Date(video.created_time)
      ]);
      
      sheet.getRange(currentRow, 1, data.length, 5).setValues(data);
      
      currentRow += data.length;
      processedCount += videos.length;
      page++;
      
      Logger.log(`Processed ${processedCount} videos from "${folder.name}" (page ${page - 1})`);
      
      // Stop if we've processed enough videos
      if (processedCount >= totalVideos) {
        break;
      }
      
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
        `https://player.vimeo.com/video/${video.uri.split('/').pop()}?`,
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

// ==================== THUMBNAIL UPLOAD FUNCTIONS ====================

// Main function to upload thumbnails
function uploadThumbnailsFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  
  // Check if we have data in the sheet
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('No Data', 'Please run "Extract from Specific Folder" first to populate the sheet with videos.', ui.ButtonSet.OK);
    return;
  }
  
  // Verify Drive folder is set
  if (!DRIVE_FOLDER_ID || DRIVE_FOLDER_ID === 'YOUR_GOOGLE_DRIVE_FOLDER_ID_HERE') {
    ui.alert('Setup Required', 
      'Please set your Google Drive folder ID in the script.\n\n' +
      'Instructions:\n' +
      '1. Upload all thumbnails to a Google Drive folder\n' +
      '2. Open that folder in Drive\n' +
      '3. Copy the folder ID from the URL (the long string after /folders/)\n' +
      '4. Paste it in line 6 of the script (DRIVE_FOLDER_ID)',
      ui.ButtonSet.OK);
    return;
  }
  
  try {
    // Get all thumbnails from Drive folder
    const thumbnails = getThumbnailsFromDrive();
    
    if (thumbnails.length === 0) {
      ui.alert('No Thumbnails', 'No image files found in the specified Google Drive folder.', ui.ButtonSet.OK);
      return;
    }
    
    const response = ui.alert(
      'Ready to Upload',
      `Found ${thumbnails.length} thumbnail(s) in Drive folder.\n` +
      `Found ${lastRow - 1} video(s) in sheet.\n\n` +
      `The script will match thumbnails to videos by name.\n\n` +
      `Continue?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // Add status column if not exists
    if (sheet.getRange(1, 6).getValue() !== 'Thumbnail Status') {
      sheet.getRange(1, 6).setValue('Thumbnail Status').setFontWeight('bold');
    }
    
    let successCount = 0;
    let errorCount = 0;
    
    // Process each video
    for (let row = 2; row <= lastRow; row++) {
      const videoTitle = sheet.getRange(row, 2).getValue(); // Column B: Video Title
      const videoId = sheet.getRange(row, 3).getValue(); // Column C: Video ID
      
      if (!videoTitle || !videoId) continue;
      
      // Find matching thumbnail
      const thumbnail = findMatchingThumbnail(videoTitle, thumbnails);
      
      if (!thumbnail) {
        sheet.getRange(row, 6).setValue('❌ No matching thumbnail');
        errorCount++;
        continue;
      }
      
      // Upload thumbnail
      sheet.getRange(row, 6).setValue('⏳ Uploading...');
      SpreadsheetApp.flush(); // Update display immediately
      
      try {
        const result = uploadThumbnailToVimeo(videoId, thumbnail);
        
        if (result.success) {
          sheet.getRange(row, 6).setValue('✅ Uploaded');
          successCount++;
        } else {
          const errorMsg = result.error || 'Upload failed';
          sheet.getRange(row, 6).setValue('❌ ' + errorMsg.substring(0, 50));
          errorCount++;
          Logger.log(`Failed for ${videoTitle} (${videoId}): ${errorMsg}`);
        }
      } catch (error) {
        sheet.getRange(row, 6).setValue('❌ Error: ' + error.message.substring(0, 40));
        errorCount++;
        Logger.log(`Exception for ${videoTitle}: ${error.message}`);
      }
      
      // Small delay to avoid rate limiting
      Utilities.sleep(1000);
    }
    
    ui.alert('Complete!', 
      `Thumbnail upload complete:\n\n` +
      `✅ Success: ${successCount}\n` +
      `❌ Failed: ${errorCount}`,
      ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Error', error.message, ui.ButtonSet.OK);
    Logger.log('Error: ' + error);
  }
}

// Get all image files from Google Drive folder
function getThumbnailsFromDrive() {
  try {
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const files = folder.getFiles();
    const thumbnails = [];
    
    while (files.hasNext()) {
      const file = files.next();
      const mimeType = file.getMimeType();
      
      // Only get image files
      if (mimeType.startsWith('image/')) {
        thumbnails.push({
          name: file.getName(),
          nameWithoutExt: getNameWithoutExtension(file.getName()),
          file: file,
          mimeType: mimeType
        });
      }
    }
    
    return thumbnails;
  } catch (error) {
    throw new Error('Could not access Drive folder. Check the folder ID and permissions.');
  }
}

// Remove file extension from filename
function getNameWithoutExtension(filename) {
  return filename.replace(/\.[^/.]+$/, '');
}

// Find thumbnail that matches video name
function findMatchingThumbnail(videoName, thumbnails) {
  // Remove common video extensions if present
  const cleanVideoName = videoName.replace(/\.(mp4|mov|avi|mkv)$/i, '');
  
  // Try exact match first
  let match = thumbnails.find(t => t.nameWithoutExt === cleanVideoName);
  if (match) return match;
  
  // Try case-insensitive match
  match = thumbnails.find(t => 
    t.nameWithoutExt.toLowerCase() === cleanVideoName.toLowerCase()
  );
  if (match) return match;
  
  // Try partial match
  match = thumbnails.find(t => 
    t.nameWithoutExt.toLowerCase().includes(cleanVideoName.toLowerCase()) ||
    cleanVideoName.toLowerCase().includes(t.nameWithoutExt.toLowerCase())
  );
  
  return match;
}

// Upload thumbnail to Vimeo video
function uploadThumbnailToVimeo(videoId, thumbnail) {
  try {
    Logger.log(`Starting upload for video ${videoId} with thumbnail ${thumbnail.name}`);
    
    // Get the image data
    const imageBlob = thumbnail.file.getBlob();
    const imageBytes = imageBlob.getBytes();
    
    Logger.log(`Image size: ${imageBytes.length} bytes, type: ${thumbnail.mimeType}`);
    
    // Step 1: Get existing pictures to see if we need to delete old ones
    const listEndpoint = `https://api.vimeo.com/videos/${videoId}/pictures`;
    
    const listOptions = {
      method: 'get',
      headers: {
        'Authorization': `bearer ${VIMEO_ACCESS_TOKEN}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };
    
    const listResponse = UrlFetchApp.fetch(listEndpoint, listOptions);
    Logger.log(`List pictures: Status ${listResponse.getResponseCode()}`);
    
    // Step 2: Create a new picture with time parameter to get upload link
    const createEndpoint = `https://api.vimeo.com/videos/${videoId}/pictures?time=0`;
    
    const createOptions = {
      method: 'post',
      headers: {
        'Authorization': `bearer ${VIMEO_ACCESS_TOKEN}`,
        'Content-Type': 'application/json',
        'Accept': 'application/vnd.vimeo.*+json;version=3.4'
      },
      muteHttpExceptions: true
    };
    
    const createResponse = UrlFetchApp.fetch(createEndpoint, createOptions);
    const createStatus = createResponse.getResponseCode();
    const createText = createResponse.getContentText();
    
    Logger.log(`Create picture: Status ${createStatus}`);
    Logger.log(`Response: ${createText.substring(0, 500)}`);
    
    if (createStatus !== 201 && createStatus !== 200) {
      return { success: false, error: `Create failed (${createStatus})` };
    }
    
    const createResult = JSON.parse(createText);
    
    // Check for upload link in the response
    let uploadLink = null;
    let pictureUri = createResult.uri;
    
    if (createResult.link) {
      // Use the direct link provided
      uploadLink = createResult.link;
    } else if (createResult.upload && createResult.upload.link) {
      uploadLink = createResult.upload.link;
    }
    
    Logger.log(`Upload link: ${uploadLink}`);
    Logger.log(`Picture URI: ${pictureUri}`);
    
    if (!uploadLink) {
      // If no upload link, try using the picture URI directly with PATCH
      const patchEndpoint = `https://api.vimeo.com${pictureUri}`;
      
      // Convert image to base64
      const base64Image = Utilities.base64Encode(imageBytes);
      
      const patchOptions = {
        method: 'patch',
        headers: {
          'Authorization': `bearer ${VIMEO_ACCESS_TOKEN}`,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify({
          active: true
        }),
        muteHttpExceptions: true
      };
      
      const patchResponse = UrlFetchApp.fetch(patchEndpoint, patchOptions);
      Logger.log(`Patch response: ${patchResponse.getResponseCode()}`);
      
      return { success: false, error: 'No upload method available' };
    }
    
    // Step 3: Upload the image to the provided link
    const uploadOptions = {
      method: 'put',
      headers: {
        'Content-Type': thumbnail.mimeType
      },
      payload: imageBytes,
      muteHttpExceptions: true
    };
    
    const uploadResponse = UrlFetchApp.fetch(uploadLink, uploadOptions);
    const uploadStatus = uploadResponse.getResponseCode();
    
    Logger.log(`Upload image: Status ${uploadStatus}`);
    
    if (uploadStatus !== 200 && uploadStatus !== 201) {
      const errorText = uploadResponse.getContentText();
      Logger.log(`Upload failed: ${errorText}`);
      return { success: false, error: `Upload failed (${uploadStatus})` };
    }
    
    // Step 4: Set as active thumbnail
    const activateEndpoint = `https://api.vimeo.com${pictureUri}`;
    
    const activateOptions = {
      method: 'patch',
      headers: {
        'Authorization': `bearer ${VIMEO_ACCESS_TOKEN}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ active: true }),
      muteHttpExceptions: true
    };
    
    const activateResponse = UrlFetchApp.fetch(activateEndpoint, activateOptions);
    const activateStatus = activateResponse.getResponseCode();
    
    Logger.log(`Activate thumbnail: Status ${activateStatus}`);
    
    Logger.log(`Success for video ${videoId}`);
    return { success: true };
    
  } catch (error) {
    Logger.log(`Exception: ${error.message}`);
    return { success: false, error: error.message };
  }
}

// ==================== MENU & HELP ====================

// Create custom menu
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Vimeo Tools')
    .addItem('📁 Extract from Specific Folder', 'extractFromSelectedFolder')
    .addItem('📋 Extract All Videos', 'extractAllVideos')
    .addSeparator()
    .addItem('🖼️ Upload Thumbnails from Drive', 'uploadThumbnailsFromSheet')
    .addSeparator()
    .addItem('ℹ️ Help', 'showHelp')
    .addToUi();
}

// Help function
function showHelp() {
  const ui = SpreadsheetApp.getUi();
  const helpText = `Vimeo Video Link Extractor & Thumbnail Uploader

SETUP:
1. Get API token from: https://developer.vimeo.com/apps
   - Needs these scopes: Public, Private, Edit, Upload
2. Paste token in line 5 (VIMEO_ACCESS_TOKEN)
3. Upload thumbnails to a Google Drive folder
4. Get folder ID from Drive URL and paste in line 6 (DRIVE_FOLDER_ID)

EXTRACT VIDEOS:
• "Extract from Specific Folder" - Select folder by name
• "Extract All Videos" - Get all videos from account

UPLOAD THUMBNAILS:
• "Upload Thumbnails from Drive" - Automatically matches 
  thumbnail names to video names and uploads them
• Thumbnails must have same name as videos (any image format)
• Example: Video "Not_Ai_1" → Thumbnail "Not_Ai_1.jpg"

Tips:
- Script matches names intelligently (case-insensitive)
- Supports: .jpg, .png, .gif, and other image formats
- Status column shows upload progress`;
  
  ui.alert('Help', helpText, ui.ButtonSet.OK);
}
