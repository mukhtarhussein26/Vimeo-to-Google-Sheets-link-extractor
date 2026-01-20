# Vimeo-to-Google-Sheets-extractor
This tool extracts  videos source links from a given library or a specific folder in Vimeo platform, and pastes it to a Google sheets document. This tool that connects to the Vimeo API to fetch video source links, IDs, and upload dates directly into your Google Spreadsheet.  


## 🚀 Features
- **Folder Selection:** Choose a specific Vimeo project/folder to extract from.
- **Bulk Extraction:** Extract every video from your entire Vimeo library at once.
- **Custom Menu:** Adds a "Vimeo Tools" menu directly to your Google Sheets UI for easy access.
- **Automated Formatting:** Automatically sets up headers, bolds them, and adjusts column widths.
- **Data Extracted:** Folder Name, Video Title, Video ID, Player URL, and Upload Date.

## 🛠️ Setup Instructions

### 1. Get a Vimeo Access Token
1. Go to the [Vimeo Developer Portal](https://developer.vimeo.com/apps).
2. Create a new app.
3. Generate a **Personal Access Token** with `Public` and `Private` scopes.

### 2. Set up the Google Sheet
1. Create a new [Google Sheet](https://sheets.new).
2. Go to **Extensions** > **Apps Script**.
3. Delete any code in the editor and paste the code from `Code.gs` in this repository.
4. Replace `YOUR_VIMEO_ACCESS_TOKEN_HERE` (line 5) with the token you generated in Step 1.
5. Save the project (click the disk icon) and name it "Vimeo Extractor".

### 3. Run the Script
1. Refresh your Google Sheet.
2. A new menu item **"Vimeo Tools"** will appear in the top toolbar.
3. Click it and select **"Extract from Specific Folder"** or **"Extract All Videos"**.
4. Authorize the script when prompted by Google.

## 📸 Preview
The script creates a custom menu like this:
`Vimeo Tools > 📁 Extract from Specific Folder`

## ⚠️ Important Note
Keep your Access Token private. If you share your spreadsheet with others, they will be able to see your token in the Apps Script editor.

## 📝 License
MIT License - Feel free to use and modify!
