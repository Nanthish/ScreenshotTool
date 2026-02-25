# SnipIT 2.0 - Windows Floating Screenshot Tool

A standalone Windows application for taking full-screen and partial screenshots with advanced markup and highlighting capabilities, then exporting them to a Word document.

## New Features (v2.0)

### 1. **Close Button (X)** 
- Added a red X button to safely close the tool without saving
- Confirms before closing to prevent accidental loss of screenshots

### 2. **Partial Screenshot Capture with Markup & Highlighting**
- **Capture Part of Screen**: Click the crosshair button to select any rectangular area of the screen
- **Live Drawing Tools**: While viewing the partial capture, choose from multiple markup tools:
  - **Highlight**: Draw semi-transparent rectangles to highlight important areas
  - **Draw**: Free-form drawing for annotations
  - **Arrow**: Draw directional arrows to point out specific elements
  - **Circle**: Draw circles around important content
  - **Rectangle**: Draw precise rectangular outlines

- **Color Selection**: Multiple colors available for markup (Red, Yellow, Green, Blue, Black, White)
- **Comments**: Add detailed comments to each partial capture

### 3. **Smart Dropdown Capture**
- While in the partial capture mode, press **'D'** to intelligently capture the currently focused window/dropdown
- Automatically detects the active window and captures it with surrounding context
- Useful for capturing dropdowns, dialogs, and menus without manual selection
- Includes padding to show context around the focused element

## Original Features

- Floating, always-on-top window with minimal UI
- Take full screenshots of the entire screen
- Add optional comments to each screenshot
- Automatic timestamp recording
- Export all screenshots to a Word document (.docx)
- No external dependencies required (standalone executable)
- Draggable window interface

## Installation

### Option 1: Using the Pre-built Executable

Download the `SnipIT.exe` file and run it directly. No installation required.

### Option 2: From Source

1. Install Python 3.8 or later
2. Install required packages:
   ```
   pip install -r requirements.txt
   ```
3. Run the application:
   ```
   python main.py
   ```

### Option 3: Create Your Own Executable

1. Install cx_Freeze:
   ```
   pip install cx_Freeze
   ```
2. Build the executable:
   ```
   python setup.py build
   ```
3. The executable will be created in the `build` directory

## Usage

### Full Screen Capture
1. Click the camera icon to capture the entire screen
2. Optionally enable the "Comment" checkbox to add notes
3. A comment dialog will appear (if enabled)

### Partial Screen Capture with Markup
1. Click the crosshair icon to start partial capture
2. Click and drag to select the area you want to capture
3. The markup window will open showing your selection
4. Choose your markup tools and colors from the controls:
   - Select a tool (Highlight, Draw, Arrow, Circle, Rectangle)
   - Select a color
   - Click and drag on the image to apply the markup
5. Click "Save with Comment" to add notes, or "Save without Comment" to save directly
6. Click "Cancel" to discard the capture

### Smart Dropdown Capture
1. Click the crosshair icon to start partial capture
2. Move the focus to the dropdown or window you want to capture
3. Press **'D'** on your keyboard
4. The tool will automatically capture the focused window with context
5. Use the markup tools to highlight important elements
6. Save with or without comments

### Export to Word
1. After capturing all screenshots/partial captures, click the green document icon
2. A Word document will open with all your captures
3. The document includes:
   - Professional report header with timestamp
   - Each capture numbered and labeled (Full or Partial)
   - Capture region coordinates for partial captures
   - Timestamps for each capture
   - Comments (if added)
   - High-quality embedded images

### Close the Tool
1. Click the red X button to close the tool
2. Confirm the action (you'll be prompted to prevent accidental closure)
3. All unsaved screenshots will be lost

## Keyboard Shortcuts

| Key | Action |
|-----|--------|
| **D** | Smart dropdown/focused window capture (during partial capture mode) |
| **ESC** | Cancel partial capture selection |
| **Click & Drag** | Select area for partial capture |

## System Requirements

- Windows 7 or later
- Python 3.8+ (for source installation)
- Microsoft Word (optional, for viewing exported documents)

## Tips

- **For Dropdowns**: Use the 'D' key for smart capture to avoid manually selecting dropdown areas
- **For Comments**: Enable the "Comment" checkbox for full-screen captures if you want to add notes
- **For Annotations**: Use different colors in markup mode to distinguish between different types of annotations
- **Dragging**: Click and drag the title bar to move the floating window around the screen
- **Non-Intrusive**: The floating window stays on top but doesn't interfere with your work

## Troubleshooting

- **Word not opening**: Make sure Microsoft Word is installed on your system
- **Capture quality**: Screenshots are saved at full screen resolution
- **Permission issues**: Run the application as Administrator if you encounter permission errors

## Future Enhancements

- Crop and rotate captured images
- Add text labels directly in markup mode
- Support for multiple monitors
- OCR text extraction from screenshots
- Cloud backup options
6. Click "End" when finished
7. A Word document will be created with all screenshots, comments, and timestamps
8. The document will open automatically

## Controls

- **Screenshot Button**: Capture the entire screen
- **End Button**: Create Word document and exit
- **Drag Window**: Click and drag the window to move it around

## Requirements

- Windows 7 or later
- Microsoft Word (for document creation and viewing)
- Python 3.8+ (if running from source)

## Technical Details

- Built with Python and Tkinter for the GUI
- Uses PIL (Pillow) for screenshot capture
- Uses python-docx for Word document generation
- Creates standalone executable with cx_Freeze
- No external dependencies in the final executable

## Troubleshooting

### Application won't start
- Make sure you have Windows 7 or later
- If running from source, ensure all requirements are installed

### Screenshots not working
- Ensure you have permission to capture screen content
- Try running as administrator

### Word document not opening
- Make sure Microsoft Word is installed
- Check that the .docx file extension is associated with Word

## License

This project is open source and available under the MIT License.

## Support

For issues and support, please create an issue on the project repository.