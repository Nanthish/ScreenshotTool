# SnipIT - Windows Floating Screenshot Tool

A standalone Windows application for taking screenshots with comments and timestamps, then exporting them to a Word document.

## Features

- Floating, always-on-top window with minimal UI
- Take screenshots of the entire screen
- Add optional comments to each screenshot
- Automatic timestamp recording
- Export all screenshots to a Word document (.docx)
- No external dependencies required (standalone executable)

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

1. Run the application
2. A floating window will appear in the top-right corner
3. Click "Screenshot" to capture the screen
4. Enter an optional comment when prompted
5. Repeat steps 3-4 as needed
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