# GamingLog

A Windows background service that automatically tracks your gaming sessions and logs them to Google Sheets.

## Features

- 🎮 **Automatic Game Detection** - Monitors Steam library for running games
- 📊 **Google Sheets Integration** - Logs sessions with game name, duration, start/end times
- 🔒 **Memory-Based Filtering** - Only tracks processes >2GB to avoid launcher duplicates
- 🚀 **Auto-Start on Login** - Runs silently in the background via Windows Startup folder
- 📝 **Detailed Logging** - Debug logs stored in `dist/gaminglog.log`

## Requirements

- Windows 10/11
- Python 3.13+
- Steam installation (for auto-detection)
- Google Cloud service account with Sheets API access

## Installation

### 1. Clone the Repository

```bash
git clone https://github.com/jaipkapoor99/GamingLog.git
cd GamingLog
```

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

### 3. Set Up Google Sheets API

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the **Google Sheets API**
4. Create a **Service Account**:
   - Navigate to **IAM & Admin** → **Service Accounts**
   - Click **Create Service Account**
   - Give it a name (e.g., "GamingLog Service")
   - Click **Create and Continue**
   - Skip role assignment (click **Continue** then **Done**)
5. Generate credentials:
   - Click on the created service account
   - Go to **Keys** tab → **Add Key** → **Create New Key**
   - Choose **JSON** format
   - Download the file and save it as `credentials.json` in the project root
6. Create a Google Sheet:

   - Create a new Google Sheet
   - Copy the **Sheet ID** from the URL:

     ```text
     https://docs.google.com/spreadsheets/d/[SHEET_ID]/edit
     ```

   - Share the sheet with the service account email (found in `credentials.json` under `client_email`)
   - Give it **Editor** permissions

### 4. Configure Environment Variables

Create a `.env` file in the project root:

```env
# Google Service Account credentials JSON file path
GOOGLE_SERVICE_ACCOUNT_FILE=C:\Coding\GamingLog\credentials.json

# Google Sheet ID (from the sheet URL)
GOOGLE_SHEET_ID=your_sheet_id_here

# Optional: Poll interval in seconds (default: 5)
POLL_INTERVAL_SECONDS=5
```

### 5. Build the Executable

```bash
python -m PyInstaller --clean GamingLog.spec
```

This creates `dist/GamingLog.exe`.

### 6. Prepare the Distribution Folder

Copy required files to the `dist` folder:

```powershell
# Copy environment configuration
Copy-Item .env .\dist\.env -Force

# Copy credentials file
Copy-Item credentials.json .\dist\credentials.json -Force
```

**Important:** Update the `.env` file in the `dist` folder to use the local credentials path:

```env
GOOGLE_SERVICE_ACCOUNT_FILE=C:\Coding\GamingLog\dist\credentials.json
GOOGLE_SHEET_ID=your_sheet_id_here
POLL_INTERVAL_SECONDS=5
```

### 7. Run GamingLog

Simply run the executable:

```powershell
.\dist\GamingLog.exe
```

**That's it!** The program will:

- ✅ Automatically install itself to Windows Startup folder on first run
- ✅ Start monitoring in the background
- ✅ Begin logging gaming sessions to Google Sheets

The executable auto-installs to:

```text
C:\Users\[USERNAME]\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\GamingLogWatcher.bat
```

## Usage

### Quick Start

Just run the executable - it handles everything automatically:

```powershell
.\dist\GamingLog.exe
```

On first run, it will install itself to Windows Startup and begin monitoring.

### Manual Commands

```bash
# Install auto-start on login (optional - happens automatically)
.\dist\GamingLog.exe --install-task

# Uninstall auto-start
.\dist\GamingLog.exe --uninstall-task

# Start watcher manually in background
.\dist\GamingLog.exe --run
```

### Google Sheet Format

The watcher creates a "Gaming" sheet with these columns:

| Game              | Duration (min) | Started At          | Ended At            |
| ----------------- | -------------- | ------------------- | ------------------- |
| Black Myth Wukong | 45.32          | 2025-11-03T16:30:00 | 2025-11-03T17:15:19 |

- **Game** - Detected game name (derived from folder name)
- **Duration (min)** - Session length in minutes (rounded to 2 decimals)
- **Started At** - ISO 8601 timestamp when game was detected
- **Ended At** - ISO 8601 timestamp when game process ended

You can customize the headers in Google Sheets - the program will continue writing to the same column positions.

### Checking Status

```powershell
# Check if watcher is running
Get-Process -Name GamingLog

# View recent logs
Get-Content .\dist\gaminglog.log -Tail 50
```

## How It Works

1. **Process Monitoring** - Polls running processes every 5 seconds using `psutil`
2. **Steam Detection** - Reads Steam library paths from Windows Registry
3. **Memory Filtering** - Only tracks game processes using >2GB RAM to avoid launcher processes
4. **Session Tracking** - Records start time when game is detected, calculates duration on exit
5. **Google Sheets Logging** - Appends session data via Google Sheets API v4

## Configuration Files

- **`.env`** - Environment variables (credentials path, sheet ID, poll interval)
- **`credentials.json`** - Google Cloud service account credentials
- **`GamingLog.spec`** - PyInstaller build configuration
- **`requirements.txt`** - Python dependencies

## Troubleshooting

### Process Not Starting

Check if already running:

```powershell
Get-Process -Name GamingLog
```

Kill existing process:

```powershell
taskkill /F /IM GamingLog.exe
```

### No Sessions Being Logged

1. Check the log file: `dist/gaminglog.log`
2. Verify game process uses >2GB memory (adjust `MIN_MEMORY_BYTES` in code if needed)
3. Ensure game is installed in Steam library folder
4. Verify Google Sheets API permissions

### Duplicate Sessions

The watcher filters processes by memory size (>2GB). If you still see duplicates, the game may be launching multiple >2GB processes. Consider adjusting the memory threshold.

### Google Sheets 404 Error

- Verify the Sheet ID is correct
- Ensure the service account email has Editor access to the sheet
- Check that `credentials.json` path is correct in `.env`

## Development

### Project Structure

```text
GamingLog/
├── main.py              # Main application logic
├── GamingLog.spec       # PyInstaller spec file
├── requirements.txt     # Python dependencies
├── .env                 # Environment configuration
├── credentials.json     # Google service account (gitignored)
├── dist/
│   ├── GamingLog.exe   # Compiled executable
│   ├── .env            # Environment file (for exe)
│   └── gaminglog.log   # Runtime logs
└── build/              # PyInstaller build artifacts
```

### Rebuilding After Changes

1. Update `main.py`
2. Kill running processes: `taskkill /F /IM GamingLog.exe`
3. Rebuild: `python -m PyInstaller --clean GamingLog.spec`
4. Copy required files to dist:

   ```powershell
   Copy-Item .env .\dist\.env -Force
   Copy-Item credentials.json .\dist\credentials.json -Force
   ```

5. Update `.env` in dist folder to point to local credentials:

   ```env
   GOOGLE_SERVICE_ACCOUNT_FILE=C:\Coding\GamingLog\dist\credentials.json
   ```

6. Restart watcher: `.\dist\GamingLog.exe`

## License

MIT

## Credits

Built with:

- [psutil](https://github.com/giampaolo/psutil) - Process monitoring
- [google-api-python-client](https://github.com/googleapis/google-api-python-client) - Google Sheets API
- [PyInstaller](https://github.com/pyinstaller/pyinstaller) - Executable packaging
- [python-dotenv](https://github.com/theskumar/python-dotenv) - Environment configuration
