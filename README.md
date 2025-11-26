# 📅 TA Roster Parser

Automated Technical Assistant (TA) duty roster downloader and JSON parser. Downloads Excel rosters from SharePoint, parses them into structured JSON format, and maintains a history of all changes (Upto a specified Time - 2 months).

## 🚀 Features

- **Automated Downloads**: Fetches the latest roster from SharePoint using Selenium WebDriver
- **Smart Parsing**: Converts Excel roster sheets to structured JSON format
- **Historical Tracking**: Maintains timestamped copies of all roster files and JSON outputs
- **Scheduled Updates**: GitHub Actions workflow runs weekly to keep roster up-to-date
- **Flexible Usage**: Can skip download and parse existing files locally
- **Auto Cleanup**: Removes roster files older than 60 days to save space

## 📥 Download Latest Roster

**Latest Roster JSON**: [Download Roster.json](https://github.com/lrxrn/TA-Roster/raw/main/roster/Roster.json)

**Latest Excel File**: [Download Roster.xlsx](https://github.com/lrxrn/TA-Roster/raw/main/roster/Roster.xlsx)

## 📊 JSON Format

The roster is parsed into the following JSON structure:

```json
[
  {
    "week": "24-11-2025",
    "parsedTimestamp": "2025-11-26T17:53:14.299Z",
    "data": [
      {
        "day": "Monday",
        "date": "24-11-2025",
        "shifts": [
          {
            "name": "S1",
            "time": {
              "start": "08:15",
              "end": "10:30"
            },
            "assignments": [
              {
                "type": "TechCentre-1",
                "person": "TA One",
                "assignType": "requested"
              },
              {
                "type": "Rounding-1",
                "person": "TA Two",
                "assignType": "assigned"
              }
            ]
          }
        ]
      }
    ]
  }
]
```

### Structure Breakdown

| Field | Type | Description |
|-------|------|-------------|
| `week` | `string` | Monday date of the roster week (DD-MM-YYYY format) |
| `parsedTimestamp` | `string` | ISO 8601 timestamp when roster was parsed (UTC) |
| `data` | `array` | Array of daily rosters (Monday-Saturday) |
| `data[].day` | `string` | Day of the week |
| `data[].date` | `string` | Specific date for that day (DD-MM-YYYY format) |
| `data[].shifts` | `array` | Array of shifts for that day |
| `data[].shifts[].name` | `string` | Shift identifier (S1, S2, S3, etc.) |
| `data[].shifts[].time` | `object` | Shift timing |
| `data[].shifts[].time.start` | `string` | Shift start time (HH:MM format) |
| `data[].shifts[].time.end` | `string` | Shift end time (HH:MM format) |
| `data[].shifts[].assignments` | `array` | Staff assignments for the shift |
| `data[].shifts[].assignments[].type` | `string` | Assignment type (TechCentre-1, Rounding-1, QC-1, etc.) |
| `data[].shifts[].assignments[].person` | `string` | Staff member name |
| `data[].shifts[].assignments[].assignType` | `string` | Either `"requested"` or `"assigned"` (yellow highlight in Excel) |

### Assignment Types

The roster includes the following shift types (May change as per roster requirements):
- **TechCentre-1** / **TechCentre-2**: Technical Centre desk assignments
- **Rounding-1** / **Rounding-2** / **Rounding-3**: Campus rounding duties
- **QC-1** / **QC-2**: Quality Control assignments

### Assignment Status

- **`requested`**: Requested shift
- **`assigned`**: Assigned shift (yellow highlighted in Excel)

## 🛠️ Installation

```bash
# Clone the repository
git clone https://github.com/lrxrn/TA-Roster.git
cd TA-Roster

# Install dependencies
npm install
```

## 📖 Usage

### Download and Parse New Roster

Downloads the latest roster from SharePoint and parses it to JSON:

```bash
node download-parse.js
```

### Parse Existing Roster (Skip Download)

Uses the most recent Excel file in the `roster/` folder:

```bash
node download-parse.js --skip-download
# or
node download-parse.js -s
```

## ⚙️ Configuration

Edit `config.json` to customize behavior:

```json
{
  "roster": {
    "downloadUrl": "https://your-sharepoint-url",
    "outputFileName": "Roster.xlsx",
    "folderName": "roster",
    "historyFolderName": "history"
  },
  "cleanup": {
    "retentionDays": 60
  },
  "download": {
    "timeoutMs": 120000,
    "headless": true
  }
}
```

## 📁 File Structure

```
TA-Roster/
├── roster/
│   ├── Roster.xlsx           # Latest Excel roster
│   ├── Roster.json           # Latest JSON roster
│   └── history/
│       ├── Roster_DD-MM-YYYY-HH-MM.xlsx
│       └── Roster_DD-MM-YYYY-HH-MM.json
├── download-parse.js         # Main script
├── config.json              # Configuration
└── package.json             # Dependencies
```

## 🤖 Automated Updates

GitHub Actions workflow runs automatically:
- **Schedule**: Every Sunday at 6:00 PM (MYT)
- **Trigger**: Can be manually triggered from Actions tab

The workflow:
1. Downloads the latest roster
2. Parses it to JSON
3. Commits changes to the repository
4. Cleans up old files (>60 days)

## 📦 Dependencies

- **selenium-webdriver**: Browser automation for downloading
- **xlsx**: Excel file parsing

## 📝 License

MIT License - See LICENSE file for details

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

---

**Note**: This tool is designed for APU Technical Assistant duty roster management. The Excel roster must follow the standard TA roster format with date headers and shift tables.
