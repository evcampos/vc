# VC Email Organization & Reporting System

This repository contains tools for organizing emails and generating Brazil VC reports.

## Features

### 📧 Email Organization

Automatically organize your inbox by moving non-priority messages to a "To be organized" folder. This helps you:
- Keep your inbox clean and focused on priority messages
- Understand email clustering and organization patterns
- Prepare messages for further processing and analysis

**Key Features:**
- ✅ Automatically creates "To be organized" folder/label if it doesn't exist
- ✅ Moves non-priority messages without deleting or replying
- ✅ Preserves priority messages (important, starred, urgent)
- ✅ Customizable priority keywords
- ✅ Dry-run mode for testing
- ✅ Batch processing with configurable limits

### 📊 Brazil VC Report Generator

Generates weekly reports on Brazil's venture capital landscape.

## Setup

### Prerequisites

- Python 3.9 or higher
- Gmail account (for email organization)
- Google Cloud project with Gmail API enabled

### Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd vc
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

### Gmail API Setup (for Email Organization)

1. **Create a Google Cloud Project:**
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project or select existing one

2. **Enable Gmail API:**
   - Navigate to "APIs & Services" > "Library"
   - Search for "Gmail API"
   - Click "Enable"

3. **Create OAuth 2.0 Credentials:**
   - Go to "APIs & Services" > "Credentials"
   - Click "Create Credentials" > "OAuth client ID"
   - Choose "Desktop app" as application type
   - Download the credentials as `credentials.json`
   - Place `credentials.json` in the root directory

## Usage

### Email Organization

#### Local Usage

Run the email organizer script:

```bash
cd scripts
python organize_emails.py
```

**Options:**

- `--dry-run`: Test mode, doesn't move messages (recommended for first run)
- `--max=N`: Process maximum N messages (default: 100)

**Examples:**

```bash
# Dry run to see what would be moved
python organize_emails.py --dry-run

# Process up to 50 messages
python organize_emails.py --max=50

# Dry run with 200 messages limit
python organize_emails.py --dry-run --max=200
```

#### GitHub Actions Usage

The workflow runs automatically every 6 hours, or you can trigger it manually:

1. Go to "Actions" tab in GitHub
2. Select "Organize Emails" workflow
3. Click "Run workflow"
4. Choose options:
   - **Dry run**: Test without making changes
   - **Max messages**: Number of messages to process

**Required Secrets:**

Configure these in GitHub Settings > Secrets:

- `GMAIL_CREDENTIALS`: Content of your `credentials.json` file
- `GMAIL_TOKEN`: (Optional) Saved authentication token from previous run

### How It Works

#### Priority Detection

Messages are considered **priority** and NOT moved if they:

1. Are marked as "Important" by Gmail
2. Are starred
3. Contain priority keywords in subject or sender:
   - urgent, important, asap, priority, critical
   - meeting, invoice, payment, contract, signed

All other messages are moved to "To be organized" folder.

#### Folder Management

The script automatically:
1. Checks if "To be organized" label exists
2. Creates it if missing (with proper visibility settings)
3. Moves messages by:
   - Adding the "To be organized" label
   - Removing the "INBOX" label
   - **NOT deleting or replying to any message**

#### Understanding the Clustering Process

The email organization tool helps you understand message clustering by:

1. **Separation**: Non-priority messages are separated from priority ones
2. **Categorization**: Messages are grouped under a common label
3. **Preservation**: All original messages remain intact for analysis
4. **Pattern Recognition**: You can analyze what types of messages end up in "To be organized"

This is useful for:
- Training machine learning models for email classification
- Understanding email patterns in your inbox
- Building custom email routing rules
- Preparing data for further clustering algorithms

## Customization

### Modify Priority Keywords

Edit `scripts/organize_emails.py` and update the `PRIORITY_KEYWORDS` list:

```python
PRIORITY_KEYWORDS = [
    'urgent', 'important', 'asap', 'priority', 'critical',
    'meeting', 'invoice', 'payment', 'contract', 'signed',
    # Add your custom keywords here
    'your-keyword', 'another-keyword'
]
```

### Change Folder Name

Edit the `ORGANIZE_FOLDER` constant:

```python
ORGANIZE_FOLDER = "Your Custom Folder Name"
```

## Security Notes

- **Never commit** `credentials.json` or `token.pickle` to version control
- These files are automatically ignored via `.gitignore`
- Store credentials securely using GitHub Secrets for automation
- The script only requests `gmail.modify` scope (read and modify messages, no deletion)

## Troubleshooting

### First Run Authentication

On first run, the script will:
1. Open a browser window
2. Ask you to log in to your Gmail account
3. Request permission to modify messages
4. Save authentication token for future runs

### Common Issues

**Error: credentials.json not found**
- Download OAuth credentials from Google Cloud Console
- Place the file in the repository root

**Error: Invalid credentials**
- Delete `token.pickle`
- Run the script again to re-authenticate

**Messages not moving**
- Check if they match priority criteria
- Use `--dry-run` to see what would be moved
- Verify Gmail API is enabled in Google Cloud Console

## Project Structure

```
.
├── .github/
│   └── workflows/
│       ├── brazil-vc-report.yml      # VC report automation
│       └── organize-emails.yml       # Email organization automation
├── scripts/
│   ├── organize_emails.py            # Email organization script
│   └── generate_report.py            # VC report generator (TBD)
├── requirements.txt                   # Python dependencies
├── README.md                          # This file
└── .gitignore                        # Ignored files
```

## Contributing

Contributions are welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## License

[Add your license here]

## Support

For issues or questions:
- Open an issue in the repository
- Check existing documentation
- Review Google Gmail API documentation
