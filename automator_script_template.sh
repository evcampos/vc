#!/bin/bash
# Automator Workflow Script Template
# Copy this entire script into Automator's "Run Shell Script" action

# ============================================================
# CONFIGURATION - UPDATE THESE PATHS
# ============================================================

# Path to this repository (update this!)
REPO_PATH="$HOME/path/to/email-clustering-system"

# Path to Python 3 (check with: which python3)
PYTHON3="/usr/local/bin/python3"

# Number of emails to process
EMAIL_LIMIT=50

# Optional: Custom database location
# DATABASE_PATH="$HOME/Dropbox/EmailClusterDatabase.xlsx"

# ============================================================
# SCRIPT - Don't modify below unless you know what you're doing
# ============================================================

# Change to repository directory
cd "$REPO_PATH" || exit 1

# Run the email clusterer
if [ -z "$DATABASE_PATH" ]; then
    # Use default database location
    "$PYTHON3" email_clusterer.py --limit "$EMAIL_LIMIT"
else
    # Use custom database location
    "$PYTHON3" email_clusterer.py --limit "$EMAIL_LIMIT" --database "$DATABASE_PATH"
fi

# Capture exit code
EXIT_CODE=$?

# Display notification
if [ $EXIT_CODE -eq 0 ]; then
    osascript -e 'display notification "Email clustering completed successfully!" with title "Email Clusterer" sound name "Glass"'
else
    osascript -e 'display notification "Email clustering failed. Check the logs." with title "Email Clusterer" sound name "Basso"'
fi

# Optional: Open the database after processing
# open "$HOME/Documents/EmailClusterDatabase.xlsx"

exit $EXIT_CODE
