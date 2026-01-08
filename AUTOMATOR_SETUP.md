# Automator Workflow Setup Guide

This guide will help you create an Automator workflow to run the Email Clustering System automatically on macOS.

## Method 1: Automator Application (Recommended)

### Step 1: Create Automator Application

1. Open **Automator** (Applications > Automator)
2. Click **New Document**
3. Select **Application** and click **Choose**

### Step 2: Add Run Shell Script Action

1. In the search box (top left), type "shell"
2. Double-click **Run Shell Script** to add it to your workflow
3. Configure the action:
   - **Shell**: `/bin/bash`
   - **Pass input**: `as arguments`

### Step 3: Add the Script

Replace the default `cat` text with:

```bash
#!/bin/bash

# Path to your email clusterer script
SCRIPT_PATH="$HOME/path/to/email_clusterer.py"

# Path to Python 3
PYTHON3="/usr/local/bin/python3"

# Alternatively, use python3 from PATH
# PYTHON3="python3"

# Run the email clusterer
$PYTHON3 "$SCRIPT_PATH" --limit 50

# Optional: Display notification when complete
osascript -e 'display notification "Email clustering complete!" with title "Email Clusterer"'
```

**Important**: Update `SCRIPT_PATH` with the actual path to `email_clusterer.py`

### Step 4: Save the Application

1. Click **File > Save** (⌘S)
2. Name it: `Email Clusterer`
3. Save location: `~/Applications` or `Desktop`
4. File Format: **Application**

### Step 5: Test the Application

1. Double-click the saved application
2. Grant permissions when prompted:
   - Allow access to Mail
   - Allow access to Calendar
   - Allow access to Documents folder

## Method 2: Automator Calendar Alarm

Run the email clusterer automatically at specific times.

### Step 1: Create Calendar Alarm

1. Open **Automator**
2. Create new **Calendar Alarm**
3. Add **Run Shell Script** action (same as above)

### Step 2: Add to Calendar

1. Save the Calendar Alarm
2. Open **Calendar** app
3. Create a new event (e.g., "Process Emails")
4. Set time: 9:00 AM daily (or your preference)
5. Click **Alert** dropdown
6. Select **Custom...**
7. Choose your saved Automator alarm
8. Set to repeat daily

## Method 3: LaunchAgent (Advanced - Runs in Background)

For fully automated background processing:

### Step 1: Create LaunchAgent plist

Create file at `~/Library/LaunchAgents/com.emailclusterer.plist`:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.emailclusterer</string>

    <key>ProgramArguments</key>
    <array>
        <string>/usr/local/bin/python3</string>
        <string>/Users/YOUR_USERNAME/path/to/email_clusterer.py</string>
        <string>--limit</string>
        <string>50</string>
    </array>

    <key>StartCalendarInterval</key>
    <dict>
        <key>Hour</key>
        <integer>9</integer>
        <key>Minute</key>
        <integer>0</integer>
    </dict>

    <key>StandardOutPath</key>
    <string>/tmp/emailclusterer.log</string>

    <key>StandardErrorPath</key>
    <string>/tmp/emailclusterer.error.log</string>
</dict>
</plist>
```

### Step 2: Load LaunchAgent

```bash
launchctl load ~/Library/LaunchAgents/com.emailclusterer.plist
```

### Step 3: Unload (if needed)

```bash
launchctl unload ~/Library/LaunchAgents/com.emailclusterer.plist
```

## Permissions Required

The script needs access to:

1. **Mail**: System Settings > Privacy & Security > Automation > [Your App/Terminal] > Mail
2. **Calendar**: System Settings > Privacy & Security > Automation > [Your App/Terminal] > Calendar
3. **Files and Folders**: Access to Documents folder for the Excel database

## Troubleshooting

### Permission Denied Errors

If you get permission errors:

1. Go to **System Settings > Privacy & Security**
2. Navigate to **Automation**
3. Enable checkboxes for Mail and Calendar for your application/Terminal

### Script Not Running

1. Check the path to `email_clusterer.py` is correct
2. Ensure Python 3 path is correct: `which python3`
3. Check logs in `/tmp/emailclusterer.log`

### Excel File Issues

If the Excel database isn't being created:
1. Ensure you have write permissions to `~/Documents`
2. Check that pandas and openpyxl are installed: `pip3 list | grep -E "pandas|openpyxl"`

## Customization Options

You can customize the script execution by modifying the command:

```bash
# Process only 20 emails
python3 email_clusterer.py --limit 20

# Use custom database location
python3 email_clusterer.py --database "/path/to/custom/database.xlsx"

# Combine options
python3 email_clusterer.py --limit 30 --database "/path/to/db.xlsx"
```

## Recommended Schedule

- **Morning**: 9:00 AM - Process overnight emails
- **Midday**: 1:00 PM - Process morning emails
- **Evening**: 5:00 PM - End of day processing

Set up multiple Calendar Alarm events for this schedule.
