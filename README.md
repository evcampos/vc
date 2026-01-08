# Email Clustering System for macOS

Automatically cluster and categorize your Apple Mail inbox by themes, with intelligent calendar event matching and a learning database system.

## Features

- **Automatic Email Clustering**: Categorizes emails into customizable themes (Work, Personal, Finance, Shopping, etc.)
- **Calendar Integration**: Detects if email subjects relate to your upcoming calendar events
- **Learning Database**: Excel-based database that you can modify and improve over time
- **Automator Compatible**: Run manually or schedule automatic processing
- **Daily Statistics**: Track email processing metrics over time
- **Detailed Logging**: Every email is logged with category, confidence score, and calendar matches

## System Requirements

- macOS 10.14 (Mojave) or later
- Python 3.7 or later
- Apple Mail
- Apple Calendar
- Microsoft Excel or compatible spreadsheet software (for viewing/editing the database)

## Quick Start

### 1. Installation

```bash
# Clone or download this repository
cd /path/to/email-clustering-system

# Run setup script
chmod +x setup.sh
./setup.sh
```

This will:
- Check Python 3 installation
- Install required packages (pandas, openpyxl)
- Make the script executable

### 2. Grant Permissions

For the script to work, you need to grant permissions:

1. Go to **System Settings** (or System Preferences on older macOS)
2. Navigate to **Privacy & Security > Automation**
3. Enable permissions for **Terminal** (or your Automator app) to control:
   - **Mail**
   - **Calendar**

### 3. First Run

Test the system with a small batch:

```bash
./email_clusterer.py --limit 10
```

This will:
- Create the Excel database at `~/Documents/EmailClusterDatabase.xlsx`
- Process 10 unread emails
- Show categorization results
- Check for calendar matches

### 4. View Results

Open the Excel database:

```bash
open ~/Documents/EmailClusterDatabase.xlsx
```

The database contains three sheets:
- **Categories**: Keyword mappings for each category
- **EmailLogs**: Complete log of all processed emails
- **Statistics**: Daily processing statistics

## Usage

### Basic Usage

```bash
# Process up to 50 unread emails (default)
./email_clusterer.py

# Process specific number of emails
./email_clusterer.py --limit 20

# Use custom database location
./email_clusterer.py --database "/path/to/custom/database.xlsx"
```

### Understanding Output

```
[1/4] Fetching emails from Apple Mail...
✓ Found 15 unread emails

[2/4] Fetching calendar events...
✓ Found 8 upcoming events

[3/4] Processing and categorizing emails...
------------------------------------------------------------

[1/15] Monthly Team Meeting - Agenda
    From: manager@company.com
    Category: Work (confidence: 0.67)
    📅 Calendar Match: Team Meeting

[2/15] Your Order Has Shipped
    From: orders@amazon.com
    Category: Shopping (confidence: 1.00)

...

SUMMARY
============================================================
Total Emails Processed: 15
Successfully Categorized: 13 (86.7%)
Calendar Matches Found: 3 (20.0%)

Database: /Users/username/Documents/EmailClusterDatabase.xlsx
============================================================
```

## Managing the Database

### Understanding the Categories Sheet

The `Categories` sheet has four columns:

| Category | Keyword | Active | Created |
|----------|---------|--------|---------|
| Work | meeting | TRUE | 2026-01-08 10:00:00 |
| Work | project | TRUE | 2026-01-08 10:00:00 |
| Finance | invoice | TRUE | 2026-01-08 10:00:00 |
| Shopping | order | TRUE | 2026-01-08 10:00:00 |

### Adding New Keywords

To improve categorization, add rows to the Categories sheet:

1. Open `EmailClusterDatabase.xlsx`
2. Go to **Categories** sheet
3. Add a new row:
   - **Category**: Choose existing or create new category name
   - **Keyword**: The keyword to match (lowercase)
   - **Active**: TRUE (to enable) or FALSE (to disable)
   - **Created**: Current date/time

Example additions:
```
Category: Work
Keyword: standup
Active: TRUE
Created: 2026-01-08 14:30:00

Category: Personal
Keyword: birthday
Active: TRUE
Created: 2026-01-08 14:30:00
```

### Creating New Categories

Simply use a new category name when adding keywords:

```
Category: Health
Keyword: doctor
Active: TRUE

Category: Health
Keyword: appointment
Active: TRUE

Category: Health
Keyword: prescription
Active: TRUE
```

### Disabling Keywords

To temporarily disable a keyword without deleting it, set `Active` to `FALSE`:

```
Category: Newsletter
Keyword: unsubscribe
Active: FALSE  ← This keyword will be ignored
```

### Analyzing Email Logs

The `EmailLogs` sheet shows every processed email:

| Timestamp | Subject | Sender | Category | CalendarMatch | MatchedEvent | Confidence |
|-----------|---------|--------|----------|---------------|--------------|------------|
| 2026-01-08 09:15:00 | Project Update | boss@work.com | Work | FALSE | | 0.67 |
| 2026-01-08 09:16:00 | Lunch Tomorrow? | friend@email.com | Personal | TRUE | Lunch with Sarah | 0.33 |

Use this data to:
- Identify miscategorized emails
- Find new keywords to add
- Track patterns over time

### Using Statistics

The `Statistics` sheet shows daily metrics:

| Date | TotalEmails | Categorized | WithCalendarMatch |
|------|-------------|-------------|-------------------|
| 2026-01-08 | 45 | 38 | 5 |
| 2026-01-07 | 52 | 43 | 7 |

Track your:
- Email volume trends
- Categorization accuracy improvement
- Calendar integration effectiveness

## Automation Setup

### Option 1: Manual Double-Click (Easiest)

1. Follow instructions in [AUTOMATOR_SETUP.md](AUTOMATOR_SETUP.md)
2. Create an Automator Application
3. Save to Desktop or Applications folder
4. Double-click whenever you want to process emails

### Option 2: Scheduled Automation

Set up automatic processing at specific times:

1. Create Automator Calendar Alarm (see [AUTOMATOR_SETUP.md](AUTOMATOR_SETUP.md))
2. Add to Calendar as recurring event
3. Runs automatically at scheduled times

Recommended schedule:
- **9:00 AM**: Morning email processing
- **1:00 PM**: Midday check
- **5:00 PM**: End-of-day processing

### Option 3: Background LaunchAgent (Advanced)

For fully automated background processing, see the LaunchAgent section in [AUTOMATOR_SETUP.md](AUTOMATOR_SETUP.md).

## Advanced Configuration

### Custom Database Location

If you want to store the database somewhere other than `~/Documents`:

```bash
./email_clusterer.py --database "/Users/username/Dropbox/EmailDB.xlsx"
```

### Processing Limits

Adjust the number of emails processed:

```bash
# Quick check (10 emails)
./email_clusterer.py --limit 10

# Full processing (100 emails)
./email_clusterer.py --limit 100

# Process ALL unread emails (use with caution)
./email_clusterer.py --limit 9999
```

### Calendar Lookback/Lookahead

The system checks calendar events 14 days ahead by default. To modify, edit `email_clusterer.py` line 245:

```python
events = self.get_calendar_events(days_ahead=14)  # Change this number
```

## How It Works

### Email Categorization

1. **Keyword Matching**: Each email subject and sender are analyzed for keywords
2. **Scoring**: Categories are scored based on number of keyword matches
3. **Confidence**: Higher confidence = more keyword matches
4. **Learning**: As you add keywords, categorization improves

### Calendar Matching

1. **Event Fetching**: Retrieves upcoming calendar events (next 14 days)
2. **Word Analysis**: Compares email subject words with event titles
3. **Matching**: Finds significant word overlap (2+ common words)
4. **Reporting**: Shows which calendar event matches the email

### Database Updates

1. **Categories**: Loaded at startup, used for classification
2. **Email Logs**: New entry for each processed email
3. **Statistics**: Daily summary updated after each run

## Troubleshooting

### "Permission denied" errors

**Solution**: Grant automation permissions
1. System Settings > Privacy & Security > Automation
2. Enable Mail and Calendar for Terminal/your Automator app

### "No unread emails to process"

This is normal if your inbox is empty or all emails are read. The system only processes **unread** emails.

### "ERROR: Required packages not installed"

**Solution**: Install dependencies
```bash
pip3 install pandas openpyxl
```

### Database file locked or in use

**Solution**: Close Excel before running the script

### Calendar events not found

**Solution**:
1. Ensure Calendar app is open at least once
2. Check calendar permissions
3. Verify you have upcoming events in the next 14 days

### AppleScript timeout errors

If processing many emails:
1. Reduce the limit: `--limit 25`
2. Process in smaller batches
3. Close unnecessary applications

## Customization Ideas

### Adding Industry-Specific Categories

**For Real Estate:**
```
Category: Listings
Keywords: listing, property, showing, open house

Category: Clients
Keywords: buyer, seller, closing, escrow
```

**For Healthcare:**
```
Category: Patients
Keywords: patient, appointment, consultation

Category: Insurance
Keywords: claim, authorization, coverage
```

### Sender-Based Categories

Add sender domains as keywords:

```
Category: Vendors
Keyword: @vendor.com
Active: TRUE

Category: Team
Keyword: @mycompany.com
Active: TRUE
```

### Priority Keywords

Create a high-priority category:

```
Category: Urgent
Keywords: urgent, asap, immediate, critical, emergency
```

## Best Practices

1. **Start Small**: Begin with 10-20 emails to test categories
2. **Review Logs**: Check EmailLogs sheet weekly to find miscategorizations
3. **Add Keywords Gradually**: Add 5-10 keywords per week based on actual emails
4. **Backup Database**: Keep a backup of your Excel file
5. **Regular Processing**: Run daily for best results
6. **Calendar Hygiene**: Keep calendar event titles descriptive for better matching

## Privacy & Security

- **Local Processing**: All processing happens on your Mac
- **No Cloud Services**: No data sent to external servers
- **No Email Storage**: Only subjects and senders are logged (not email bodies)
- **Secure**: Uses macOS security framework and requires explicit permissions

## File Structure

```
email-clustering-system/
├── email_clusterer.py          # Main Python script
├── requirements.txt            # Python dependencies
├── setup.sh                    # Installation script
├── AUTOMATOR_SETUP.md         # Automator workflow guide
└── README.md                  # This file

Generated files:
~/Documents/EmailClusterDatabase.xlsx  # Your email database
```

## Contributing

Found a bug or have suggestions? Feel free to:
1. Modify the script for your needs
2. Add new categories and keywords
3. Share your customizations

## License

This project is provided as-is for personal use.

## Support

For issues or questions:
1. Check the Troubleshooting section
2. Review [AUTOMATOR_SETUP.md](AUTOMATOR_SETUP.md) for automation issues
3. Verify permissions in System Settings

---

**Happy Email Organizing!** 📧✨
