# Vimeo Video Manager: Link Extractor & Thumbnail Uploader

An all in- ne Google Apps Script tool to manage Vimeo libraries via Google Sheets. This tool not only extracts video data but also automates the process of uploading thumbnails from Google Drive.

## 🚀 New Features
- **Thumbnail Automation:** Automatically matches image files in a Google Drive folder to video titles in your sheet and uploads them to Vimeo.
- **Intelligent Matching:** Matches names case-insensitively and ignores file extensions (e.g., `Video_01.mp4` matches `Video_01.jpg`).
- **Status Tracking:** Provides real-time feedback in the spreadsheet (⏳ Uploading, ✅ Success, ❌ Error).
- **Expanded Extraction:** Handles pagination better for large video libraries.

## 🛠️ Setup Instructions

### 1. Vimeo API Configuration
1. Go to the [Vimeo Developer Portal](https://developer.vimeo.com/apps).
2. **Important:** Your Access Token now needs the following scopes: `Public`, `Private`, `Edit`, and `Upload`.
3. Generate the token and keep it ready.

### 2. Google Drive Configuration
1. Create a folder in Google Drive and upload your thumbnail images.
2. Ensure the image filenames match your Vimeo video titles (e.g., Video Title: "Intro", Image: "Intro.png").
3. Copy the **Folder ID** from the URL (the string of characters after `/folders/`).

### 3. Script Installation
1. Open a Google Sheet and go to **Extensions** > **Apps Script**.
2. Paste the code from `Code.gs`.
3. Replace the placeholders at the top:
   - `VIMEO_ACCESS_TOKEN`: Your generated token.
   - `DRIVE_FOLDER_ID`: The ID of your Google Drive thumbnail folder.
4. Save and refresh your spreadsheet.

## 📁 How to Use
1. **Extract:** Click `Vimeo Tools > 📁 Extract from Specific Folder`.
2. **Review:** Ensure the "Video Title" column matches your image names in Drive.
3. **Upload:** Click `Vimeo Tools > 🖼️ Upload Thumbnails from Drive`.
4. **Monitor:** Watch the "Thumbnail Status" column for results.

## ⚠️ Security Note
Never share your Google Sheet or GitHub code with your actual Token or Drive ID included. Always use placeholders when sharing.

## 📝 License
MIT
