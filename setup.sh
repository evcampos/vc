#!/bin/bash
# Email Clustering System Setup Script for macOS

echo "======================================"
echo "Email Clustering System Setup"
echo "======================================"
echo ""

# Check if Python 3 is installed
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 is not installed."
    echo "Please install Python 3 from https://www.python.org/downloads/"
    exit 1
fi

echo "✓ Python 3 found: $(python3 --version)"
echo ""

# Check if pip is installed
if ! command -v pip3 &> /dev/null; then
    echo "❌ pip3 is not installed."
    echo "Please install pip3"
    exit 1
fi

echo "✓ pip3 found"
echo ""

# Install required packages
echo "Installing required Python packages..."
pip3 install -r requirements.txt

if [ $? -eq 0 ]; then
    echo ""
    echo "✓ All packages installed successfully!"
else
    echo ""
    echo "❌ Package installation failed. Please check the error messages above."
    exit 1
fi

# Make the main script executable
chmod +x email_clusterer.py

echo ""
echo "======================================"
echo "Setup Complete!"
echo "======================================"
echo ""
echo "Next steps:"
echo "1. Grant permissions to Terminal/Python to access:"
echo "   - Mail (System Settings > Privacy & Security > Automation)"
echo "   - Calendar (System Settings > Privacy & Security > Automation)"
echo ""
echo "2. Test the script:"
echo "   ./email_clusterer.py --limit 10"
echo ""
echo "3. Set up Automator workflow (see AUTOMATOR_SETUP.md)"
echo ""
