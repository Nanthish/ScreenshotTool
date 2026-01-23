#!/usr/bin/env python3
"""
Setup script for creating a standalone executable of SnipIT
"""

from cx_Freeze import setup, Executable
import sys
import os

# Dependencies are automatically detected, but it might need fine tuning.
build_options = {
    'packages': [],
    'excludes': [],
    'include_files': [],
    'include_msvcr': True,
}

# Base configuration for Windows
base = 'Win32GUI' if sys.platform == 'win32' else None

executables = [
    Executable(
        'main.py',
        base=base,
        target_name='SnipIT.exe',
        icon='SnipIT.ico',
        shortcut_name="SnipIT",
        shortcut_dir="DesktopFolder"
    )
]

setup(
    name='SnipIT',
    version='1.0.0',
    description='SnipIT - Windows Floating Screenshot Tool with Word Export',
    options={'build_exe': build_options},
    executables=executables
)