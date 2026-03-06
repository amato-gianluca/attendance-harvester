# Quick Start - No Azure App Needed! 🚀

Get started in 3 minutes with just your Microsoft credentials!

## Steps

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Create Configuration
```bash
cp config.yaml.template config.yaml
```

The default configuration already uses Microsoft Graph PowerShell's public client, so you don't need to change anything!

### 3. Run the Tool
```bash
python main.py
```

### 4. Authenticate
When prompted, you'll see:
```
======================================================================
AUTHENTICATION REQUIRED
======================================================================
To sign in, use a web browser to open the page https://microsoft.com/devicelogin
and enter the code XXXXXXXXX to authenticate.
======================================================================
```

1. Open https://microsoft.com/devicelogin in your browser
2. Enter the code shown
3. Login with your Microsoft/Office 365 credentials
4. That's it!

## What Happens Next?

The tool will:
1. ✅ Scan all your Teams
2. ✅ Filter teams by regex (default: all teams)
3. ✅ Find meetings from the last 30 days
4. ✅ Download attendance reports for meetings you organized
5. ✅ Save logs to `output/` folder as CSV and JSON

## Customization

Edit `config.yaml` to:

```yaml
# Filter specific teams
team_filter:
  regex: "^Project.*"  # Only teams starting with "Project"

# Change date range
meetings:
  lookback_days: 7  # Last 7 days instead of 30

# Change output format
output:
  format: "csv"  # Only CSV (or "json" or "both")
```

## Command-Line Options

```bash
# Verbose logging
python main.py -v

# Override team filter
python main.py --team-regex "Marketing|Sales"

# Different lookback period
python main.py --lookback-days 7

# Clear cache and re-authenticate
python main.py --clear-cache
```

## Troubleshooting

### No attendance data found
- You can only see attendance for meetings **you organized**
- Attendance reports are only available after meetings end
- Check that meetings are within the lookback period

### Permission errors
If you get permission errors with the default public client, you may need to:
1. Create your own Azure AD app (see full README)
2. Grant explicit permissions
3. Ask your IT admin for consent

### Works but some teams are missing
Adjust the regex filter in config.yaml:
```yaml
team_filter:
  regex: ".*"  # Match ALL teams
```

## Need More Help?

See the full [README.md](README.md) for:
- Creating your own Azure AD app
- Detailed configuration options
- Advanced troubleshooting
- Scheduling automated runs

---

**You're ready to go!** Just run `python main.py` and authenticate with your Microsoft credentials. No Azure portal needed! 🎉
