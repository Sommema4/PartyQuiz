# PartyQuiz
Scripts for automatic PartyQuiz generation from Google Sheets to Google Slides

## Setup Instructions

### 1. Enable Google APIs (One-time setup)

You already have `credentials.json`, which is great! If you need to set this up again or for reference:

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Select your project: `pubquiz-451721`
3. Click "Enable APIs and Services"
4. Enable these two APIs:
   - **Google Sheets API**
   - **Google Slides API**

### 2. Install Python Dependencies

```bash
pip install -r requirements.txt
```

### 3. Run the Script

```bash
python generate_google_slides.py
```

**First Time Running:**
- A browser window will open asking you to log in with your Google account
- Click "Allow" to grant permissions for reading spreadsheets and creating presentations
- The script will save a `token.json` file so you don't need to authenticate again

**Subsequent Runs:**
- The script will use the saved `token.json` automatically
- No browser popup needed!

## How It Works

1. **Authentication**: Uses OAuth2 to securely access your private Google Spreadsheet (no need to make it public!)
2. **Read Data**: Reads quiz questions from your spreadsheet
3. **Create Slides**: Automatically creates a Google Slides presentation with your quiz content
4. **Token Storage**: Saves authentication token so you don't need to log in every time

## Configuration

Edit these variables in `generate_google_slides.py`:

```python
SPREADSHEET_ID = "1zFjlmAF7oBPbnfDVmFPs9itCdFem4XTQaBSYM_JgfxM"  # Your spreadsheet ID
RANGE_NAME = 'otazky!A1:C100'  # The sheet name and range to read
```

## Troubleshooting

**"Access blocked: This app's request is invalid"**
- Make sure you've enabled both Google Sheets API and Google Slides API in your Google Cloud Console

**"invalid_grant" error**
- Delete `token.json` and run the script again to re-authenticate

**"The caller does not have permission"**
- Make sure you're logged in with the Google account that owns the spreadsheet

## Files

- `generate_google_slides.py` - Main script
- `credentials.json` - OAuth2 client configuration (keep this private!)
- `token.json` - Authentication token (auto-generated, keep this private!)
- `requirements.txt` - Python dependencies
