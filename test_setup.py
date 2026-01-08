#!/usr/bin/env python3
"""
Setup verification script for Email Clustering System
Tests that all dependencies and permissions are properly configured
"""

import sys
import subprocess
from pathlib import Path


def test_python_version():
    """Check Python version"""
    print("Testing Python version...", end=" ")
    version = sys.version_info
    if version.major >= 3 and version.minor >= 7:
        print(f"✓ Python {version.major}.{version.minor}.{version.micro}")
        return True
    else:
        print(f"✗ Python {version.major}.{version.minor} (need 3.7+)")
        return False


def test_dependencies():
    """Check if required packages are installed"""
    print("\nTesting dependencies:")

    packages = ['pandas', 'openpyxl']
    all_installed = True

    for package in packages:
        try:
            __import__(package)
            print(f"  ✓ {package}")
        except ImportError:
            print(f"  ✗ {package} (not installed)")
            all_installed = False

    return all_installed


def test_mail_access():
    """Test access to Apple Mail"""
    print("\nTesting Apple Mail access...", end=" ")

    applescript = '''
    tell application "Mail"
        get name of inbox
    end tell
    '''

    try:
        result = subprocess.run(
            ['osascript', '-e', applescript],
            capture_output=True,
            text=True,
            timeout=10
        )

        if result.returncode == 0:
            print("✓ Can access Mail")
            return True
        else:
            print("✗ Cannot access Mail")
            print(f"     Error: {result.stderr.strip()}")
            return False

    except Exception as e:
        print(f"✗ Error: {e}")
        return False


def test_calendar_access():
    """Test access to Apple Calendar"""
    print("Testing Apple Calendar access...", end=" ")

    applescript = '''
    tell application "Calendar"
        get name of calendars
    end tell
    '''

    try:
        result = subprocess.run(
            ['osascript', '-e', applescript],
            capture_output=True,
            text=True,
            timeout=10
        )

        if result.returncode == 0:
            print("✓ Can access Calendar")
            return True
        else:
            print("✗ Cannot access Calendar")
            print(f"     Error: {result.stderr.strip()}")
            return False

    except Exception as e:
        print(f"✗ Error: {e}")
        return False


def test_file_permissions():
    """Test write permissions to Documents folder"""
    print("Testing file permissions...", end=" ")

    try:
        test_file = Path.home() / "Documents" / ".email_clusterer_test"
        test_file.write_text("test")
        test_file.unlink()
        print("✓ Can write to Documents")
        return True
    except Exception as e:
        print(f"✗ Cannot write to Documents: {e}")
        return False


def main():
    """Run all tests"""
    print("="*60)
    print("Email Clustering System - Setup Verification")
    print("="*60)

    tests = [
        ("Python Version", test_python_version),
        ("Dependencies", test_dependencies),
        ("Mail Access", test_mail_access),
        ("Calendar Access", test_calendar_access),
        ("File Permissions", test_file_permissions),
    ]

    results = {}

    for name, test_func in tests:
        try:
            results[name] = test_func()
        except Exception as e:
            print(f"✗ {name} test failed: {e}")
            results[name] = False

    # Summary
    print("\n" + "="*60)
    print("SUMMARY")
    print("="*60)

    passed = sum(results.values())
    total = len(results)

    for name, result in results.items():
        status = "✓ PASS" if result else "✗ FAIL"
        print(f"{status}: {name}")

    print(f"\nTotal: {passed}/{total} tests passed")

    if passed == total:
        print("\n🎉 All tests passed! Your system is ready.")
        print("\nNext step: Run your first email clustering:")
        print("  ./email_clusterer.py --limit 5")
    else:
        print("\n⚠️  Some tests failed. Please fix the issues above.")
        print("\nCommon fixes:")

        if not results.get("Dependencies"):
            print("  • Install dependencies: pip3 install pandas openpyxl")

        if not results.get("Mail Access") or not results.get("Calendar Access"):
            print("  • Grant permissions:")
            print("    System Settings > Privacy & Security > Automation")
            print("    Enable Mail and Calendar for Terminal")

        if not results.get("File Permissions"):
            print("  • Check Documents folder permissions")

    print("="*60)

    return 0 if passed == total else 1


if __name__ == '__main__':
    sys.exit(main())
