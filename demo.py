#!/usr/bin/env python3
"""
Demo script for Windows Automation Hello World use cases.

This script allows you to choose which backend to use for creating
a Hello World document in your Downloads folder.
"""

import sys
import os

def show_menu():
    """Display the main menu."""
    print("=" * 60)
    print("Windows Automation Toolkit - Hello World Demo")
    print("=" * 60)
    print("Choose an automation backend:")
    print()
    print("1. Win32 API (Low-level Windows API)")
    print("   - Direct Windows API calls")
    print("   - Creates: hello_world_win32.txt")
    print()
    print("2. pywinauto (High-level automation)")
    print("   - Automated GUI interaction")
    print("   - Creates: hello_world_pywinauto.txt")
    print()
    print("3. Both (Run both demos)")
    print("4. Exit")
    print()

def run_win32_demo():
    """Run the win32 Hello World demo."""
    print("\n" + "=" * 40)
    print("Running Win32 Demo...")
    print("=" * 40)
    
    try:
        from learn_windows_automation.win32 import create_hello_world_document
        success = create_hello_world_document()
        
        if success:
            print("\n✅ Win32 demo completed successfully!")
        else:
            print("\n❌ Win32 demo failed!")
        
        return success
    except Exception as e:
        print(f"\n❌ Error running win32 demo: {e}")
        return False

def run_pywinauto_demo():
    """Run the pywinauto Hello World demo."""
    print("\n" + "=" * 40)
    print("Running pywinauto Demo...")
    print("=" * 40)
    
    try:
        from learn_windows_automation.pywinauto import create_hello_world_document
        success = create_hello_world_document()
        
        if success:
            print("\n✅ pywinauto demo completed successfully!")
        else:
            print("\n❌ pywinauto demo failed!")
            
        return success
    except Exception as e:
        print(f"\n❌ Error running pywinauto demo: {e}")
        return False

def main():
    """Main function."""
    while True:
        show_menu()
        
        try:
            choice = input("Enter your choice (1-4): ").strip()
            
            if choice == '1':
                run_win32_demo()
                input("\nPress Enter to continue...")
                
            elif choice == '2':
                run_pywinauto_demo()
                input("\nPress Enter to continue...")
                
            elif choice == '3':
                print("\nRunning both demos...")
                win32_success = run_win32_demo()
                pywinauto_success = run_pywinauto_demo()
                
                print("\n" + "=" * 50)
                print("SUMMARY:")
                print(f"Win32 demo: {'✅ Success' if win32_success else '❌ Failed'}")
                print(f"pywinauto demo: {'✅ Success' if pywinauto_success else '❌ Failed'}")
                
                downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
                print(f"\nCheck your Downloads folder: {downloads_path}")
                input("\nPress Enter to continue...")
                
            elif choice == '4':
                print("\nGoodbye!")
                break
                
            else:
                print("\n❌ Invalid choice! Please enter 1, 2, 3, or 4.")
                input("Press Enter to continue...")
                
        except KeyboardInterrupt:
            print("\n\nGoodbye!")
            break
        except Exception as e:
            print(f"\n❌ Unexpected error: {e}")
            input("Press Enter to continue...")

if __name__ == "__main__":
    main()
