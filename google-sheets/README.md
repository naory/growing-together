# Google Sheets Setup

Scripts for creating and managing Google Sheets for the Growing Together project.

## Quick Setup

### 1. Create Service Account

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a project or select existing one
3. Enable **Google Sheets API**:
   - Navigate to **APIs & Services** → **Library**
   - Search for "Google Sheets API" and enable it
4. Create Service Account:
   - Go to **APIs & Services** → **Credentials**
   - Click **Create Credentials** → **Service Account**
   - Fill in name (e.g., "growing-together-sheets") and click **Create and Continue**
   - Skip role assignment, click **Done**
5. Download Service Account Key:
   - Click on the service account email
   - Go to **Keys** tab → **Add Key** → **Create new key**
   - Select **JSON** format and download
   - Rename to `service_account.json` and place in this directory

### 2. Create and Share Spreadsheet

1. Go to [Google Sheets](https://sheets.google.com)
2. Create a new blank spreadsheet
3. Optionally rename it (e.g., "צומחים_ביחד_דאטה")
4. Click **Share** button (top right)
5. Add the service account email (from `service_account.json` → `client_email`)
   - Example: `growing-together-sheets@growingtogether-483409.iam.gserviceaccount.com`
6. Give it **Editor** permissions and click **Send**
7. Copy the Spreadsheet ID from the URL:
   - URL format: `https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit`
   - Copy the `SPREADSHEET_ID` part
8. Update `create-sheet.py`:
   ```python
   EXISTING_SPREADSHEET_ID = "your-spreadsheet-id-here"
   ```

### 3. Run the Script

```bash
# Activate virtual environment
source ../venv/bin/activate

# Run the script
python create-sheet.py
```

The script will:
- Add required sheets (if missing)
- Set up headers and picklists
- Configure data validations (dropdowns)
- Apply formatting (frozen headers, RTL, bold)

## Security

- **Never commit** `service_account.json` to version control
- Keep service account keys secure
- Rotate keys periodically
