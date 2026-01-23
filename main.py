#!/usr/bin/env python3
"""
SnipIT - Windows Floating Screenshot Tool
A standalone application for taking screenshots with comments and timestamps,
then exporting them to a Word document.
"""

import sys
import os
import time
import threading
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from PIL import Image, ImageGrab, ImageTk, ImageDraw
import pythoncom
import win32clipboard
from docx import Document
from docx.shared import Inches, RGBColor
from datetime import datetime
import tempfile
import subprocess
import ctypes
from ctypes import wintypes
import win32gui
import win32con
import pyautogui
import win32com.client
import xml.etree.ElementTree as ET
import base64


class ScreenshotTool:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("SnipIT")
        self.root.geometry("116x140")
        self.root.attributes("-topmost", True)
        self.root.attributes("-alpha", 0.9)
        
        # Make window floating and borderless but keep it in taskbar
        self.root.overrideredirect(True)
        
        # Set window style to make it always on top and show in taskbar
        hwnd = self.root.winfo_id()
        ctypes.windll.user32.SetWindowLongW(hwnd, -20, 0x00000008 | 0x00000080 | 0x00000020 | 0x00040000)
        
        # Initialize data storage
        self.screenshots = []
        self.is_capturing = False
        self.add_comment_var = tk.BooleanVar(value=False)
        
        self.setup_ui()
        self.center_window()
        
    def create_button_icons(self):
        """Create icon images for buttons using PIL"""
        # Screenshot icon (camera icon - 40x40)
        screenshot_icon = Image.new("RGBA", (40, 40), (255, 255, 255, 0))
        draw = ImageDraw.Draw(screenshot_icon)
        # Draw camera body
        draw.rectangle([8, 12, 32, 28], fill="#2196F3", outline="#1976D2", width=2)
        # Draw lens
        draw.ellipse([14, 16, 26, 28], fill="#64B5F6", outline="#1976D2", width=2)
        # Draw flash
        draw.rectangle([28, 8, 32, 12], fill="#FFC107", outline="#FFA000", width=1)
        self.screenshot_icon = ImageTk.PhotoImage(screenshot_icon)
        
        # End icon (document icon - 40x40)
        end_icon = Image.new("RGBA", (40, 40), (255, 255, 255, 0))
        draw = ImageDraw.Draw(end_icon)
        # Draw document
        draw.rectangle([10, 8, 30, 32], fill="#4CAF50", outline="#388E3C", width=2)
        # Draw lines on document
        draw.line([14, 14, 26, 14], fill="white", width=2)
        draw.line([14, 18, 26, 18], fill="white", width=2)
        draw.line([14, 22, 26, 22], fill="white", width=2)
        draw.line([14, 26, 22, 26], fill="white", width=2)
        self.end_icon = ImageTk.PhotoImage(end_icon)
    
    def setup_ui(self):
        """Setup the floating UI with buttons"""
        self.create_button_icons()
        
        # Main frame
        main_frame = tk.Frame(self.root, bg="#f0f0f0", padx=3, pady=3)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Bind drag to main frame
        main_frame.bind('<Button-1>', self.start_move)
        main_frame.bind('<B1-Motion>', self.do_move)
        
        # Title label - make it draggable
        self.title_label = tk.Label(main_frame, text="SnipIT", 
                              font=("Arial", 8, "bold"), bg="#f0f0f0", fg="#333333", cursor="hand2")
        self.title_label.pack(pady=(0, 2), fill=tk.X)
        
        # Bind drag events to title label
        self.title_label.bind('<Button-1>', self.start_move)
        self.title_label.bind('<B1-Motion>', self.do_move)
        
        # Button frame
        button_frame = tk.Frame(main_frame, bg="#f0f0f0")
        button_frame.pack(fill=tk.X, pady=0)
        
        # Bind drag to button frame too
        button_frame.bind('<Button-1>', self.start_move)
        button_frame.bind('<B1-Motion>', self.do_move)
        
        # Screenshot button with icon
        self.screenshot_btn = tk.Button(button_frame, image=self.screenshot_icon,
                                       command=self.take_screenshot, 
                                       bg="#ffffff", relief=tk.RAISED, bd=2,
                                       cursor="hand2")
        self.screenshot_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # End button with icon
        self.end_btn = tk.Button(button_frame, image=self.end_icon,
                                command=self.end_session,
                                bg="#ffffff", relief=tk.RAISED, bd=2,
                                cursor="hand2")
        self.end_btn.pack(side=tk.LEFT)
        
        # Checkbox for comments (below buttons)
        checkbox_frame = tk.Frame(main_frame, bg="#f0f0f0")
        checkbox_frame.pack(fill=tk.X, pady=(2, 0))
        
        self.comment_checkbox = tk.Checkbutton(checkbox_frame, text="Comment", 
                                               variable=self.add_comment_var,
                                               bg="#f0f0f0", font=("Arial", 7))
        self.comment_checkbox.pack(side=tk.LEFT)
        
        # Bind drag to checkbox frame
        checkbox_frame.bind('<Button-1>', self.start_move)
        checkbox_frame.bind('<B1-Motion>', self.do_move)
        
    def center_window(self):
        """Position the floating window 45 pixels away from both corners"""
        try:
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            
            x = 45  # 45 pixels from left edge
            y = 45  # 45 pixels from top edge
            
            self.root.geometry(f"116x140+{x}+{y}")
        except tk.TclError:
            pass
        
    def start_move(self, event):
        """Start dragging the window"""
        # Only start move if not clicking on buttons
        if event.widget.__class__.__name__ != 'Button':
            self.x = event.x_root
            self.y = event.y_root
        
    def do_move(self, event):
        """Move the window while dragging"""
        # Only move if not on buttons
        if event.widget.__class__.__name__ != 'Button':
            deltax = event.x_root - self.x
            deltay = event.y_root - self.y
            x = self.root.winfo_x() + deltax
            y = self.root.winfo_y() + deltay
            self.root.geometry(f"+{x}+{y}")
            self.x = event.x_root
            self.y = event.y_root
        
    def take_screenshot(self):
        """Take a screenshot of the entire screen"""
        if self.is_capturing:
            return
            
        self.is_capturing = True
        
        try:
            # Hide the floating window temporarily
            self.root.withdraw()
            
            # Small delay to ensure window is hidden
            self.root.update()
            time.sleep(0.1)
            
            # Take screenshot
            screenshot = ImageGrab.grab()
            
            # Show window again
            self.root.deiconify()
            
            # Get comment only if checkbox is enabled
            comment = ""
            if self.add_comment_var.get():
                comment = self.get_comment()
                if comment is None:  # User clicked cancel
                    self.is_capturing = False
                    return
                    
            # Store screenshot data
            timestamp = datetime.now()
            self.screenshots.append({
                'image': screenshot,
                'comment': comment,
                'timestamp': timestamp
            })
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to take screenshot: {str(e)}")
        finally:
            self.is_capturing = False
            
    def get_comment(self):
        """Get a comment from the user for the screenshot"""
        comment = simpledialog.askstring("Add Comment", 
                                       "Enter a comment for this screenshot (optional):",
                                       parent=self.root)
        return comment if comment is not None else ""
        
    def update_status(self, message):
        """No longer updates any status label"""
        pass
        
    def end_session(self):
        """End the session and create Word document"""
        if not self.screenshots:
            messagebox.showinfo("No Screenshots", "No screenshots were taken.")
            return
            
        try:
            # Create Word document directly
            doc = Document()
            
            # Add screenshots with comments and timestamps
            for i, shot in enumerate(self.screenshots, 1):
                # Add screenshot number as text with blue color and bold
                p = doc.add_paragraph()
                run = p.add_run(f"Screenshot {i}")
                run.bold = True
                run.font.color.rgb = RGBColor(0, 0, 255)  # Blue color
                
                # Add timestamp
                timestamp_text = shot["timestamp"].strftime("%Y-%m-%d %H:%M:%S")
                doc.add_paragraph(f"Timestamp: {timestamp_text}")
                
                # Add comment if exists
                if shot["comment"]:
                    doc.add_paragraph(f"Comment: {shot['comment']}")
                
                # Save image to temporary file and add to document
                temp_file = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
                shot["image"].save(temp_file.name, "PNG")
                temp_file.close()
                
                # Add image to document
                doc.add_picture(temp_file.name, width=Inches(6))
                
                # Clean up temp file
                os.unlink(temp_file.name)
                
                # Add page break except for last screenshot
                if i < len(self.screenshots):
                    doc.add_page_break()
            
            # Save to a temporary file
            temp_docx = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
            temp_docx_path = temp_docx.name
            temp_docx.close()
            doc.save(temp_docx_path)
            
            # Open in Word using COM interface
            try:
                word = win32com.client.GetObject(Class="Word.Application")
            except:
                # If Word isn't already running, start it
                word = win32com.client.Dispatch("Word.Application")
            
            word.Visible = True
            
            # Open the document
            word_doc = word.Documents.Open(FileName=os.path.abspath(temp_docx_path))
            
            # Bring Word to foreground
            word.Activate()
            
            # Clean up - delete temp file after a delay to ensure Word has loaded it
            def cleanup_temp():
                time.sleep(2)
                try:
                    os.unlink(temp_docx_path)
                except:
                    pass
            
            cleanup_thread = threading.Thread(target=cleanup_temp, daemon=True)
            cleanup_thread.start()
            
            # Automatically quit the tool after a short delay
            self.root.after(500, self.root.quit)
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create document: {str(e)}")
            
    def run(self):
        """Start the application"""
        self.root.mainloop()


def main():
    """Main entry point"""
    app = ScreenshotTool()
    app.run()


if __name__ == "__main__":
    main()
