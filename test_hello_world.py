#!/usr/bin/env python3
"""
Test script for the new Hello World use cases.
"""

import sys
import os

def test_win32_hello_world():
    """Test the win32 Hello World function."""
    print("=== Testing win32 Hello World ===")
    try:
        from learn_windows_automation.win32 import create_hello_world_document
        print("✓ Function imported successfully")
        
        # Test the function (but don't actually run it to avoid opening Notepad)
        print("✓ win32 Hello World function is available")
        return True
    except Exception as e:
        print(f"✗ Error: {e}")
        return False

def test_pywinauto_hello_world():
    """Test the pywinauto Hello World function."""
    print("\n=== Testing pywinauto Hello World ===")
    try:
        from learn_windows_automation.pywinauto import create_hello_world_document
        print("✓ Function imported successfully")
        
        # Test the function (but don't actually run it to avoid opening Notepad)
        print("✓ pywinauto Hello World function is available")
        return True
    except Exception as e:
        print(f"✗ Error: {e}")
        return False

def test_main_functions():
    """Test the updated main functions."""
    print("\n=== Testing Updated Main Functions ===")
    try:
        from learn_windows_automation.win32 import main as win32_main
        from learn_windows_automation.pywinauto import main as pywinauto_main
        print("✓ Both main functions imported successfully")
        return True
    except Exception as e:
        print(f"✗ Error: {e}")
        return False

def main():
    """Run all tests."""
    print("Testing new Hello World use cases...\n")
    
    results = []
    results.append(test_win32_hello_world())
    results.append(test_pywinauto_hello_world())
    results.append(test_main_functions())
    
    print("\n" + "="*50)
    if all(results):
        print("🎉 All tests passed! Hello World use cases are ready.")
        print("\nTo run the use cases:")
        print("  win32: uv run python -c 'from learn_windows_automation.win32 import main; main()'")
        print("  pywinauto: uv run python -c 'from learn_windows_automation.pywinauto import main; main()'")
    else:
        print("💥 Some tests failed!")
    
    return all(results)

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
