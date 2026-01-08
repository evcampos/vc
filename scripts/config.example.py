"""
Example configuration file for email organization.
Copy this to config.py and customize as needed.
"""

# Folder/Label name for organizing non-priority messages
ORGANIZE_FOLDER = "To be organized"

# Priority keywords - messages containing these will NOT be moved
# Add or remove keywords based on your needs
PRIORITY_KEYWORDS = [
    # Urgency indicators
    'urgent',
    'important',
    'asap',
    'priority',
    'critical',
    'emergency',

    # Business related
    'meeting',
    'invoice',
    'payment',
    'contract',
    'signed',
    'deadline',
    'approval',

    # Project related
    'milestone',
    'release',
    'launch',

    # Add your custom keywords below
    # 'custom-keyword-1',
    # 'custom-keyword-2',
]

# Maximum number of messages to process in one run
MAX_MESSAGES_DEFAULT = 100

# Enable verbose logging
VERBOSE = True
