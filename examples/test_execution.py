#!/usr/bin/env python3
"""Test script to verify script execution."""

import sys

print("Hello from test script!")
print(f"Arguments received: {sys.argv[1:]}")

if len(sys.argv) > 1:
    print(f"First argument: {sys.argv[1]}")

print("Script executed successfully!")
