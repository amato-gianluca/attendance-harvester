# Microsoft Teams Attendance Harvester

A Python tool to automatically scan Microsoft Teams, filter teams by name, discover meetings, download attendance logs, and export them to CSV/JSON files.

## Features

- 🔐 **Secure Authentication**: Uses Microsoft MSAL with device code flow and token caching
- 🔍 **Team Filtering**: Filter teams using regular expressions
- 📅 **Meeting Discovery**: Automatically discovers Teams meetings from your calendar
- 📊 **Attendance Export**: Downloads attendance reports and exports to CSV and/or JSON
- 💾 **Checkpoint System**: Avoids reprocessing already-downloaded attendance data
- 🔄 **Retry Logic**: Handles API throttling and transient errors automatically
- ⚙️ **Configurable**: YAML-based configuration for all settings

## Prerequisites

- Python 3.8 or higher
- Microsoft Azure AD application with appropriate permissions (see Setup section)
- Access to Microsoft Teams with organizer rights for meetings you want to track

## Installation

1. **Clone or download this repository**

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure**:
   ```bash
   cp config.yaml.template config.yaml
   ```

   The default configuration uses Microsoft Graph PowerShell's public client ID, so **you don't need to create an Azure app**! Just use your Microsoft credentials when prompted.

## Two Ways to Use This Tool

### Option 1: Simple Setup (Recommended) - No Azure App Needed! ✨

**Just login with your Microsoft credentials - no setup required!**

The tool comes pre-configured with Microsoft Graph PowerShell's public client ID (`14d82eec-204b-4c2f-b7e8-296a70dab67e`), which is already registered by Microsoft. You can use it immediately:

1. Keep the default `client_id` in [config.yaml](config.yaml):
   ```yaml
   azure:
     client_id: "14d82eec-204b-4c2f-b7e8-296a70dab67e"  # Microsoft Graph PowerShell
     tenant_id: "common"
     authority: "https://login.microsoftonline.com/common"
   ```

2. Run the tool:
   ```bash
   python main.py
   ```

3. Authenticate with your Microsoft credentials in the browser when prompted

**That's it!** No Azure portal, no app registration, no admin permissions needed.

**Limitations of this approach:**
- You can only request permissions that have already been consented to for Microsoft Graph PowerShell
- Some organizations may restrict access to well-known public clients
- If you get permission errors, you may need to use Option 2

### Option 2: Custom Azure App (Full Control)

If you need specific permissions or Option 1 doesn't work for your organization, you can create your own Azure AD application:

#### Step 1: Register Application

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations** → **New registration**
3. Configure:
   - **Name**: Teams Attendance Harvester (or any name you prefer)
   - **Supported account types**: "Accounts in this organizational directory only" (single tenant)
   - **Redirect URI**: Leave blank for now
4. Click **Register**

### Step 2: Note Application Details

After registration, note down:
- **Application (client) ID**
- **Directory (tenant) ID**

### Step 3: Enable Device Code Flow

1. In your app registration, go to **Authentication**
2. Under **Advanced settings** → **Allow public client flows**, set to **Yes**
3. Click **Save**

### Step 4: Grant API Permissions

1. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**
2. Add the following permissions:
   - `Team.ReadBasic.All`
   - `Channel.ReadBasic.All`
   - `Calendars.Read`
   - `OnlineMeetings.Read`
   - `OnlineMeetingArtifact.Read.All`
3. Click **Add permissions**
4. **(Optional but recommended)** Click **Grant admin consent** if you have admin rights

#### Step 5: Update Configuration

Replace the `client_id` in [config.yaml](config.yaml) with your own:

```yaml
azure:
  client_id: "YOUR_CLIENT_ID_HERE"          # Your custom client ID from Azure Portal
  tenant_id: "YOUR_TENANT_ID_HERE"          # Your tenant ID or "common"
  authority: "https://login.microsoftonline.com/YOUR_TENANT_ID_HERE"
```

## Qu# Default: Microsoft Graph PowerShell client (no Azure app needed!)
     client_id: "14d82eec-204b-4c2f-b7e8-296a70dab67e"
     tenant_id: "common"
     authority: "https://login.microsoftonline.com/common"

     # OR use your own Azure app:
     # client_id: "YOUR_CLIENT_ID_HERE"
     # tenant_id: "YOUR_TENANT_ID_HERE"
     #*Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **Copy and edit config** (the default works out of the box!):
   ```bash
   cp config.yaml.template config.yaml
   # Edit if needed, but default uses Microsoft Graph PowerShell client
   ```

3. **Run**:
   ```bash
   python main.py
   ```

4. **Authenticate in browser** when prompted with your Microsoft credentials

## Configuration

1. **Copy the template configuration**:
   ```bash
   cp config.yaml.template config.yaml
   ```

2. **Edit config.yaml** with your settings:

   ```yaml
   azure:
     client_id: "YOUR_CLIENT_ID_HERE"          # From Azure Portal
     tenant_id: "YOUR_TENANT_ID_HERE"          # From Azure Portal
     authority: "https://login.microsoftonline.com/YOUR_TENANT_ID_HERE"

   team_filter:
     regex: ".*"                                # Regex to filter team names

   meetings:
     lookback_days: 30                          # Days to look back for meetings
     include_associated_teams: true             # Include shared channel teams
     general_channel_only: true                 # Only scan General channel

   output:
     directory: "./output"                      # Output directory
     format: "both"                             # "csv", "json", or "both"
   ```

### Team Filter Examples

```yaml
# Match all teams
regex: ".*"

# Teams starting with "Project"
regex: "^Project.*"

# Teams containing "Marketing" or "Sales"
regex: "(Marketing|Sales)"

# Teams with year patterns (e.g., "Team 2024", "Project 2025")
regex: ".*202[4-6].*"
```

## Usage

### Basic Usage

Run the harvester with default settings from config.yaml:

```bash
python main.py
```

### Command-Line Options

```bash
# Use custom configuration file
python main.py -c /path/to/config.yaml

# Enable verbose logging
python main.py -v

# Clear token cache and re-authenticate
python main.py --clear-cache

# Override team filter from command line
python main.py --team-regex "^Project.*"

# Override lookback period
python main.py --lookback-days 7
```

### First Run - Authentication

On first run (or after clearing cache), you'll see:

```
======================================================================
AUTHENTICATION REQUIRED
======================================================================
To sign in, use a web browser to open the page https://microsoft.com/devicelogin
and enter the code XXXXXXXXX to authenticate.
======================================================================
```

1. Open the URL in your browser
2. Enter the code displayed
3. Sign in with your Microsoft account
4. The token will be cached for future runs

## Output Files

The tool creates files in the output directory with this naming pattern:

```
{team_name}_{channel_name}_{meeting_date}_{meeting_id}_{report_id}_attendance.csv
{team_name}_{channel_name}_{meeting_date}_{meeting_id}_{report_id}_attendance.json
```

Example:
```
Marketing_Team_General_20260306_1430_a1b2c3d4_e5f6g7h8_attendance.csv
Marketing_Team_General_20260306_1430_a1b2c3d4_e5f6g7h8_attendance.json
```

### CSV Format

The CSV file contains flattened attendance records:

| email | display_name | role | total_attendance_duration | join_datetime | leave_datetime | meeting_subject | meeting_start | meeting_organizer |
|-------|--------------|------|---------------------------|---------------|----------------|-----------------|---------------|-------------------|
| ... | ... | ... | ... | ... | ... | ... | ... | ... |

### JSON Format

The JSON file contains the complete raw data from Microsoft Graph API, including:
- Meeting metadata
- Report information
- Detailed attendance records with all intervals

## Project Structure

```
camafi/
├── main.py                      # Main entry point
├── requirements.txt             # Python dependencies
├── config.yaml.template         # Configuration template
├── config.yaml                  # Your configuration (not in git)
├── README.md                    # This file
├── src/
│   ├── __init__.py
│   ├── auth.py                  # MSAL authentication
│   ├── graph_client.py          # Microsoft Graph API client
│   ├── team_filter.py           # Team filtering logic
│   ├── meeting_resolver.py      # Meeting discovery & attendance extraction
│   └── exporter.py              # CSV/JSON export logic
├── cache/                       # Token cache & checkpoints (not in git)
└── output/                      # Attendance log files (not in git)
```

## Troubleshooting

### No attendance data found

**Possible causes**:
1. **You are not the meeting organizer**: Attendance reports are only available to meeting organizers
2. **No meetings in time range**: Adjust `lookback_days` in config
3. **Attendance not yet generated**: Reports are created after meetings end
4. **Already processed**: Use `--clear-cache` or delete `cache/processed_meetings.json`

**Solution**: Ensure you organize the meetings, or run the script with the organizer's account.

### Authentication fails

**Possible causes**:
1. **Incorrect client_id or tenant_id**: Verify values from Azure Portal
2. **Public client flow not enabled**: Check Azure AD app registration settings
3. **Missing permissions**: Ensure all required API permissions are granted

### Rate limiting / 429 errors

The tool automatically handles rate limiting with exponential backoff. If you consistently hit rate limits:
- Reduce `lookback_days` to scan fewer meetings
- Increase `retry_backoff_factor` in config

### Permission errors

If you see 403/Forbidden errors:
- Verify all required permissions are granted in Azure AD
- If admin consent is required, ask your IT admin to grant consent
- Ensure you're using delegated (not application) permissions

## Important Notes

1. **Organizer-Only Access**: You can only retrieve attendance for meetings where you are the organizer
2. **Retention Limits**: Attendance data has retention limits (typically up to 1 year)
3. **Channel Meeting Limitations**: Some channel meeting types have limited attendance API support
4. **Privacy**: Handle exported attendance data according to your organization's privacy policies

## Advanced Usage

### Running on a Schedule

You can run this tool on a schedule using cron (Linux/Mac) or Task Scheduler (Windows):

**Linux/Mac (cron)**:
```bash
# Run daily at 9 AM
0 9 * * * cd /path/to/camafi && python main.py >> logs/attendance.log 2>&1
```

**Windows (Task Scheduler)**:
Create a batch file and schedule it via Task Scheduler GUI.

### Filtering by Date in Filenames

To organize outputs by date, modify the `filename_pattern` in config:

```yaml
output:
  filename_pattern: "{meeting_date}/{team_name}_{meeting_id}_{report_id}_attendance"
```

This creates date-based subdirectories in the output folder.

## Dependencies

- `msal>=1.24.0` - Microsoft Authentication Library
- `requests>=2.31.0` - HTTP library
- `pyyaml>=6.0` - YAML configuration parser
- `python-dateutil>=2.8.2` - Date/time utilities

## License

This project is provided as-is for educational and automation purposes. Ensure compliance with your organization's IT and privacy policies when using this tool.

## Support

For issues, questions, or contributions:
1. Check existing issues and documentation
2. Review Microsoft Graph API documentation
3. Verify Azure AD application configuration
4. Enable verbose logging (`-v`) to debug issues

## Contributing

Contributions are welcome! Areas for improvement:
- Better team-to-meeting association logic
- Support for app-only (daemon) authentication
- Enhanced error recovery
- Additional export formats
- Web UI for configuration

---

**Version**: 1.0.0
**Last Updated**: March 2026
