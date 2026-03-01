# SnipIT - Windows Screenshot Tool
## Executable Distribution

### What is SnipIT?
SnipIT is a floating Windows screenshot tool that allows you to capture full or partial screenshots and annotate them before exporting to Word documents.

### File Information
- **Executable**: `SnipIT.exe`
- **Size**: 25.73 MB (includes all dependencies)
- **Location**: `dist/` folder
- **Platform**: Windows 10/11

### Getting Started

1. **Run the Application**
   - Double-click `SnipIT.exe` to launch
   - A floating widget will appear in the top-left corner
   - The application works even when minimized or unfocused

2. **Keyboard Shortcuts**
   - `Ctrl+Alt+F` - Full screen capture
   - `Ctrl+Alt+P` - Partial capture (drag to select region)

### Features

✅ **Full Screenshot Capture**
- Captures entire screen with one hotkey
- Hide floating widget automatically during capture
- Clean image without UI interference

✅ **Partial Screenshot Capture**
- Drag to select any region on screen
- 5-second countdown for dropdown/menu preparation
- Visual selection rectangle with corner handles
- Escape key to cancel

✅ **Image Annotation**
- Rectangle tool
- Circle tool
- Freehand draw tool
- Color selection (Red, Yellow, Green)
- Clear button to remove all markups

✅ **Comments & Timestamps**
- Add optional comments to screenshots
- Automatic timestamp recording
- Full metadata in exported document

✅ **Word Document Export**
- Click "End" button to export all screenshots
- Automatically opens in Microsoft Word
- Professional formatting with screenshots, timestamps, and comments

### Floating Widget Controls
- **📷** - Full screenshot (Ctrl+Alt+F)
- **⊞** - Partial screenshot (Ctrl+Alt+P)
- **📄** - End session & export to Word
- **?** - Help information
- **✕** - Close application

### System Requirements
- Windows 10 or later
- Microsoft Office Word (for export feature)
- 50 MB free disk space
- 4GB RAM recommended

### Technical Details

**Included Dependencies**:
- Python 3.11
- Tkinter (GUI)
- PIL/Pillow (Image processing)
- python-docx (Word document generation)
- pyautogui (Automation)
- win32 libraries (Windows API integration)
- All necessary DLLs and modules

**Build Configuration**:
- One-file executable for easy distribution
- Windowed mode (no console window)
- All dependencies bundled

### Troubleshooting

**Issue**: Application won't start
- Solution: Ensure you have .NET Framework installed
- Try running as Administrator

**Issue**: Hotkeys not working
- Solution: Ensure no other application is using Ctrl+Alt+F/P
- Restart the application

**Issue**: Word export fails
- Solution: Install Microsoft Word
- Ensure Word is configured properly

**Issue**: Partial capture doesn't work
- Solution: Make selection at least 10x10 pixels
- Try pressing Escape and retrying

### Support & Contact

**Work Email**: nanthish.t@gds.ey.com
**Personal Email**: nanthishwaran579@gmail.com

### Version
SnipIT v2.0 - Released March 2026

### License
Internal Use Only

---

**Note**: This executable includes all dependencies and libraries. No Python installation required!
