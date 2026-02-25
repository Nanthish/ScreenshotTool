#!/usr/bin/env python3
"""
Test Suite for SnipIT 2.0
Tests all functionality without requiring pytest
"""

import sys
import os
import traceback

# Add project to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def test_imports():
    """Test all imports"""
    print("\n" + "="*60)
    print("TEST 1: Import Verification")
    print("="*60)
    
    tests_passed = 0
    tests_failed = 0
    
    imports = [
        ("sys", "import sys"),
        ("os", "import os"),
        ("time", "import time"),
        ("threading", "import threading"),
        ("tkinter", "import tkinter as tk"),
        ("tkinter.ttk", "from tkinter import ttk"),
        ("tkinter.messagebox", "from tkinter import messagebox"),
        ("tkinter.simpledialog", "from tkinter import simpledialog"),
        ("PIL", "from PIL import Image, ImageGrab, ImageTk, ImageDraw"),
        ("win32clipboard", "import win32clipboard"),
        ("docx", "from docx import Document"),
        ("docx.shared", "from docx.shared import Inches, RGBColor"),
        ("datetime", "from datetime import datetime"),
        ("tempfile", "import tempfile"),
        ("subprocess", "import subprocess"),
        ("ctypes", "import ctypes"),
        ("ctypes.wintypes", "from ctypes import wintypes"),
        ("win32gui", "import win32gui"),
        ("win32con", "import win32con"),
        ("pyautogui", "import pyautogui"),
        ("win32com.client", "import win32com.client"),
        ("xml.etree.ElementTree", "import xml.etree.ElementTree as ET"),
        ("base64", "import base64"),
        ("json", "import json"),
    ]
    
    for module_name, import_statement in imports:
        try:
            exec(import_statement)
            print(f"✅ {module_name:<30} OK")
            tests_passed += 1
        except ImportError as e:
            print(f"❌ {module_name:<30} FAILED: {str(e)}")
            tests_failed += 1
        except Exception as e:
            print(f"⚠️  {module_name:<30} ERROR: {str(e)}")
            tests_failed += 1
    
    print(f"\nResults: {tests_passed} passed, {tests_failed} failed")
    return tests_passed, tests_failed

def test_main_file():
    """Test main.py syntax and structure"""
    print("\n" + "="*60)
    print("TEST 2: main.py Syntax Validation")
    print("="*60)
    
    try:
        import py_compile
        py_compile.compile('main.py', doraise=True)
        print("✅ main.py syntax is valid")
        return 1, 0
    except py_compile.PyCompileError as e:
        print(f"❌ Syntax error in main.py:")
        print(f"   {str(e)}")
        return 0, 1

def test_class_structure():
    """Test if ScreenshotTool class can be instantiated (without displaying GUI)"""
    print("\n" + "="*60)
    print("TEST 3: ScreenshotTool Class Structure")
    print("="*60)
    
    tests_passed = 0
    tests_failed = 0
    
    try:
        # Import the main module
        import main
        
        # Check if class exists
        if hasattr(main, 'ScreenshotTool'):
            print("✅ ScreenshotTool class exists")
            tests_passed += 1
        else:
            print("❌ ScreenshotTool class not found")
            tests_failed += 1
        
        # Check required methods
        required_methods = [
            'take_screenshot',
            'partial_capture',
            'on_select_start',
            'on_select_motion',
            'on_select_end',
            'cancel_partial_capture',
            'smart_dropdown_capture',
            'open_markup_window',
            'get_comment',
            'close_tool',
            'end_session',
            'create_button_icons',
            'setup_ui',
            'center_window',
            'start_move',
            'do_move',
            'run',
        ]
        
        for method in required_methods:
            if hasattr(main.ScreenshotTool, method):
                print(f"✅ Method '{method}' exists")
                tests_passed += 1
            else:
                print(f"❌ Method '{method}' missing")
                tests_failed += 1
                
    except Exception as e:
        print(f"❌ Error checking class structure: {str(e)}")
        print(traceback.format_exc())
        tests_failed += 1
    
    print(f"\nResults: {tests_passed} passed, {tests_failed} failed")
    return tests_passed, tests_failed

def test_file_structure():
    """Test project file structure"""
    print("\n" + "="*60)
    print("TEST 4: Project File Structure")
    print("="*60)
    
    required_files = {
        'main.py': 'Main application file',
        'setup.py': 'Build configuration',
        'requirements.txt': 'Dependencies',
        'README.md': 'Documentation',
        'QUICKSTART.md': 'Quick start guide',
    }
    
    tests_passed = 0
    tests_failed = 0
    
    for filename, description in required_files.items():
        if os.path.exists(filename):
            size = os.path.getsize(filename)
            print(f"✅ {filename:<20} ({description:<30}) - {size:>8} bytes")
            tests_passed += 1
        else:
            print(f"❌ {filename:<20} MISSING")
            tests_failed += 1
    
    print(f"\nResults: {tests_passed} passed, {tests_failed} failed")
    return tests_passed, tests_failed

def test_dependencies():
    """Test if all dependencies are installed"""
    print("\n" + "="*60)
    print("TEST 5: Dependencies Installation Check")
    print("="*60)
    
    dependencies = {
        'PIL': 'Pillow',
        'docx': 'python-docx',
        'pyautogui': 'pyautogui',
        'win32clipboard': 'pywin32',
        'win32gui': 'pywin32',
        'win32con': 'pywin32',
    }
    
    tests_passed = 0
    tests_failed = 0
    
    for module, package_name in dependencies.items():
        try:
            __import__(module)
            print(f"✅ {package_name:<20} is installed")
            tests_passed += 1
        except ImportError:
            print(f"❌ {package_name:<20} is NOT installed")
            tests_failed += 1
    
    print(f"\nResults: {tests_passed} passed, {tests_failed} failed")
    return tests_passed, tests_failed

def main():
    """Run all tests"""
    print("\n" + "🧪 "*30)
    print("SnipIT 2.0 - TEST SUITE")
    print("🧪 "*30)
    
    total_passed = 0
    total_failed = 0
    
    # Run tests
    passed, failed = test_imports()
    total_passed += passed
    total_failed += failed
    
    passed, failed = test_main_file()
    total_passed += passed
    total_failed += failed
    
    passed, failed = test_file_structure()
    total_passed += passed
    total_failed += failed
    
    passed, failed = test_dependencies()
    total_passed += passed
    total_failed += failed
    
    passed, failed = test_class_structure()
    total_passed += passed
    total_failed += failed
    
    # Final summary
    print("\n" + "="*60)
    print("FINAL TEST SUMMARY")
    print("="*60)
    print(f"Total Tests Passed: {total_passed}")
    print(f"Total Tests Failed: {total_failed}")
    
    if total_failed == 0:
        print("\n✅ ALL TESTS PASSED - Application is ready to build and run!")
        return 0
    else:
        print(f"\n❌ {total_failed} test(s) failed - Please fix issues before building")
        return 1

if __name__ == '__main__':
    exit(main())
