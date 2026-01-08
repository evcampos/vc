# Quick Start Guide - 5 Minutes to Email Clustering

Follow these steps to get up and running quickly.

## Step 1: Install (2 minutes)

Open Terminal and run:

```bash
cd /path/to/this/folder
chmod +x setup.sh
./setup.sh
```

Wait for installation to complete.

## Step 2: Grant Permissions (1 minute)

1. Open **System Settings** (or System Preferences)
2. Go to **Privacy & Security** → **Automation**
3. Find **Terminal** in the list
4. Check the boxes for:
   - ✅ **Mail**
   - ✅ **Calendar**

## Step 3: Run Your First Test (1 minute)

```bash
./email_clusterer.py --limit 5
```

You should see:
- Email fetching
- Calendar event loading
- Email categorization results
- Summary statistics

## Step 4: Check the Database (1 minute)

```bash
open ~/Documents/EmailClusterDatabase.xlsx
```

You'll see three sheets:
1. **Categories** - Your keyword rules
2. **EmailLogs** - Processed email history
3. **Statistics** - Daily metrics

## Step 5: Customize (Optional)

### Add Your Own Keywords

In the Excel file, go to the **Categories** sheet and add rows:

| Category | Keyword | Active | Created |
|----------|---------|--------|---------|
| Work | standup | TRUE | 2026-01-08 15:00:00 |
| Personal | gym | TRUE | 2026-01-08 15:00:00 |
| Finance | tax | TRUE | 2026-01-08 15:00:00 |

Save and run again!

## Next Steps

### Set Up Automation

See [AUTOMATOR_SETUP.md](AUTOMATOR_SETUP.md) to:
- Create a double-click application
- Schedule automatic processing
- Set up background execution

### Run Daily

```bash
# Process all new emails (up to 50)
./email_clusterer.py
```

### Improve Over Time

1. Check **EmailLogs** weekly
2. Find miscategorized emails
3. Add relevant keywords to **Categories**
4. Your system gets smarter!

## Common First-Time Issues

### "Permission denied"
→ Complete Step 2 above

### "No module named pandas"
→ Run: `pip3 install pandas openpyxl`

### "No unread emails"
→ Mark some emails as unread to test

### "Can't find python3"
→ Install Python from https://www.python.org/downloads/

## Tips for Success

1. **Start small**: Process 5-10 emails first
2. **Review results**: Check if categories make sense
3. **Add keywords**: Based on your actual email patterns
4. **Run daily**: Consistency improves accuracy
5. **Backup**: Keep a copy of your Excel database

## What's Being Analyzed?

✅ Email subjects
✅ Sender addresses
✅ Calendar event titles
❌ Email bodies (not read for privacy)
❌ Attachments (not accessed)

## Need Help?

Check the full [README.md](README.md) for:
- Detailed documentation
- Troubleshooting guide
- Advanced configuration
- Customization examples

---

**You're all set!** Your emails will now be automatically categorized and matched with your calendar. 🎉
