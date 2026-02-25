#!/usr/bin/env python3
"""
SnipIT 2.0 - Windows Floating Screenshot Tool
Enhanced with partial capture, markup, highlighting, and comments.
A standalone application for taking full and partial screenshots with markup,
then exporting them to a Word document.
"""

import sys
import os
import time
import threading
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
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
import json


class ScreenshotTool:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("SnipIT 2.0")
        self.root.geometry("160x140")
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
        self.partial_cap_window = None
        self.start_x = 0
        self.start_y = 0
        self.current_x = 0
        self.current_y = 0
        self.is_selecting = False
        
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
        
        # Partial capture icon (crosshair - 40x40)
        partial_icon = Image.new("RGBA", (40, 40), (255, 255, 255, 0))
        draw = ImageDraw.Draw(partial_icon)
        # Draw crosshair
        draw.line([20, 5, 20, 15], fill="#FF9800", width=2)
        draw.line([20, 25, 20, 35], fill="#FF9800", width=2)
        draw.line([5, 20, 15, 20], fill="#FF9800", width=2)
        draw.line([25, 20, 35, 20], fill="#FF9800", width=2)
        # Draw center circle
        draw.ellipse([17, 17, 23, 23], fill="#FF9800", outline="#F57C00", width=1)
        self.partial_icon = ImageTk.PhotoImage(partial_icon)
        
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
        
        # Close icon (X - 40x40)
        close_icon = Image.new("RGBA", (40, 40), (255, 255, 255, 0))
        draw = ImageDraw.Draw(close_icon)
        # Draw X
        draw.line([12, 12, 28, 28], fill="#F44336", width=3)
        draw.line([28, 12, 12, 28], fill="#F44336", width=3)
        self.close_icon = ImageTk.PhotoImage(close_icon)
    
    def setup_ui(self):
        """Setup the floating UI with buttons"""
        self.create_button_icons()
        
        # Main frame
        main_frame = tk.Frame(self.root, bg="#f0f0f0", padx=3, pady=3)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Bind drag to main frame
        main_frame.bind('<Button-1>', self.start_move)
        main_frame.bind('<B1-Motion>', self.do_move)
        
        # Top frame with title and close button
        top_frame = tk.Frame(main_frame, bg="#f0f0f0")
        top_frame.pack(fill=tk.X, pady=(0, 2))
        
        # Title label - make it draggable
        self.title_label = tk.Label(top_frame, text="SnipIT 2.0", 
                              font=("Arial", 8, "bold"), bg="#f0f0f0", fg="#333333", cursor="hand2")
        self.title_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Bind drag events to title label
        self.title_label.bind('<Button-1>', self.start_move)
        self.title_label.bind('<B1-Motion>', self.do_move)
        
        # Close button in top right corner (small red button with white X)
        self.close_btn = tk.Button(top_frame, text="✕", command=self.close_tool,
                                  bg="#E63946", fg="white", relief=tk.FLAT, bd=0,
                                  cursor="hand2", font=("Arial", 10, "bold"),
                                  padx=3, pady=0, width=2, height=1)
        self.close_btn.pack(side=tk.RIGHT, padx=(3, 0))
        
        # Bind drag to top frame
        top_frame.bind('<Button-1>', self.start_move)
        top_frame.bind('<B1-Motion>', self.do_move)
        
        # Button frame 1 (first row)
        button_frame1 = tk.Frame(main_frame, bg="#f0f0f0")
        button_frame1.pack(fill=tk.X, pady=0)
        
        # Bind drag to button frame
        button_frame1.bind('<Button-1>', self.start_move)
        button_frame1.bind('<B1-Motion>', self.do_move)
        
        # Screenshot button with icon
        self.screenshot_btn = tk.Button(button_frame1, image=self.screenshot_icon,
                                       command=self.take_screenshot, 
                                       bg="#ffffff", relief=tk.RAISED, bd=2,
                                       cursor="hand2")
        self.screenshot_btn.pack(side=tk.LEFT, padx=(0, 3))
        
        # Partial capture button
        self.partial_btn = tk.Button(button_frame1, image=self.partial_icon,
                                    command=self.partial_capture,
                                    bg="#ffffff", relief=tk.RAISED, bd=2,
                                    cursor="hand2")
        self.partial_btn.pack(side=tk.LEFT, padx=(0, 3))
        
        # End button with icon
        self.end_btn = tk.Button(button_frame1, image=self.end_icon,
                                command=self.end_session,
                                bg="#ffffff", relief=tk.RAISED, bd=2,
                                cursor="hand2")
        self.end_btn.pack(side=tk.LEFT, padx=(0, 3))
        
        # Checkbox for comments (below buttons)
        checkbox_frame = tk.Frame(main_frame, bg="#f0f0f0")
        checkbox_frame.pack(fill=tk.X, pady=(2, 0))
        
        self.comment_checkbox = tk.Checkbutton(checkbox_frame, text="Comment", 
                                               variable=self.add_comment_var,
                                               bg="#f0f0f0", font=("Arial", 7))
        self.comment_checkbox.pack(side=tk.LEFT)
        
        # Help button in right corner (blue button with white ?)
        self.help_btn = tk.Button(checkbox_frame, text="?", command=self.show_help,
                                  bg="#1976D2", fg="white", relief=tk.FLAT, bd=0,
                                  cursor="hand2", font=("Arial", 9, "bold"),
                                  padx=3, pady=0, width=2, height=1)
        self.help_btn.pack(side=tk.RIGHT, padx=(3, 0))
        
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
            
            self.root.geometry(f"160x140+{x}+{y}")
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
                'timestamp': timestamp,
                'markup_data': None
            })
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to take screenshot: {str(e)}")
        finally:
            self.is_capturing = False
    
    def partial_capture(self):
        """Start partial screen capture with selection"""
        if self.is_capturing:
            return
        
        self.is_capturing = True
        self.root.withdraw()
        self.root.update()
        time.sleep(0.2)
        
        try:
            # Create full screen overlay for selection
            self.partial_cap_window = tk.Toplevel()
            self.partial_cap_window.attributes("-fullscreen", True)
            self.partial_cap_window.attributes("-alpha", 0.1)
            self.partial_cap_window.configure(bg="gray")
            self.partial_cap_window.attributes("-topmost", True)
            
            # Instructions label
            info_label = tk.Label(self.partial_cap_window, 
                                 text="Click and drag to select area | Press 'D' for smart dropdown capture | Press ESC to cancel",
                                 bg="black", fg="yellow", font=("Arial", 10))
            info_label.pack(side=tk.TOP, fill=tk.X)
            
            # Create canvas for drawing selection rectangle
            self.canvas = tk.Canvas(self.partial_cap_window, cursor="crosshair", 
                                   bg="gray", highlightthickness=0)
            self.canvas.pack(fill=tk.BOTH, expand=True)
            
            # Bind mouse events for selection
            self.canvas.bind("<Button-1>", self.on_select_start)
            self.canvas.bind("<B1-Motion>", self.on_select_motion)
            self.canvas.bind("<ButtonRelease-1>", self.on_select_end)
            self.canvas.bind("<Escape>", self.cancel_partial_capture)
            self.canvas.bind("<d>", self.smart_dropdown_capture)
            self.canvas.bind("<D>", self.smart_dropdown_capture)
            
            self.is_selecting = False
            self.selection_rect = None
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to start capture: {str(e)}")
            self.is_capturing = False
            if self.partial_cap_window:
                self.partial_cap_window.destroy()
    
    def on_select_start(self, event):
        """Start selection"""
        self.is_selecting = True
        self.start_x = event.x
        self.start_y = event.y
        
        # Delete previous rectangle if exists
        if self.selection_rect:
            self.canvas.delete(self.selection_rect)
    
    def on_select_motion(self, event):
        """Update selection rectangle while dragging"""
        if not self.is_selecting:
            return
        
        self.current_x = event.x
        self.current_y = event.y
        
        # Delete previous rectangle
        if self.selection_rect:
            self.canvas.delete(self.selection_rect)
        
        # Draw new rectangle
        self.selection_rect = self.canvas.create_rectangle(
            self.start_x, self.start_y, self.current_x, self.current_y,
            outline="#C0C0C0", width=3
        )
    
    def on_select_end(self, event):
        """Finish selection and capture"""
        self.is_selecting = False
        
        if self.selection_rect:
            self.canvas.delete(self.selection_rect)
        
        # Get coordinates
        x1 = min(self.start_x, self.current_x)
        y1 = min(self.start_y, self.current_y)
        x2 = max(self.start_x, self.current_x)
        y2 = max(self.start_y, self.current_y)
        
        # Validate selection
        if x2 - x1 < 10 or y2 - y1 < 10:
            messagebox.showwarning("Invalid Selection", 
                                 "Please select a larger area (minimum 10x10 pixels)")
            self.canvas.create_text(self.canvas.winfo_width()//2, 
                                   self.canvas.winfo_height()//2,
                                   text="Click and drag to select area\nPress ESC to cancel",
                                   fill="white", font=("Arial", 14))
            return
        
        # Close overlay
        self.partial_cap_window.destroy()
        
        # Capture the selected area
        try:
            screenshot = ImageGrab.grab(bbox=(x1, y1, x2, y2))
            
            # Show main window
            self.root.deiconify()
            
            # Open markup window
            self.open_markup_window(screenshot, x1, y1, x2, y2)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to capture: {str(e)}")
            self.root.deiconify()
        finally:
            self.is_capturing = False
    
    def cancel_partial_capture(self, event):
        """Cancel partial capture"""
        self.is_selecting = False
        if self.partial_cap_window:
            self.partial_cap_window.destroy()
        self.root.deiconify()
        self.is_capturing = False
    
    def smart_dropdown_capture(self, event):
        """Smart capture of focused dropdown or element with surrounding context"""
        try:
            # Close the overlay
            if self.partial_cap_window:
                self.partial_cap_window.destroy()
            
            # Get the window with focus
            focused_hwnd = win32gui.GetForegroundWindow()
            
            # Get window dimensions
            rect = win32gui.GetWindowRect(focused_hwnd)
            x1, y1, x2, y2 = rect
            
            # Expand the capture area slightly to include dropdown context
            # Add some padding to capture the dropdown and surrounding elements
            padding = 50
            x1 = max(0, x1 - padding)
            y1 = max(0, y1 - padding)
            x2 = min(win32gui.GetSystemMetrics(0), x2 + padding)
            y2 = min(win32gui.GetSystemMetrics(1), y2 + padding)
            
            # Capture the area
            screenshot = ImageGrab.grab(bbox=(x1, y1, x2, y2))
            
            # Show main window
            self.root.deiconify()
            
            # Open markup window
            self.open_markup_window(screenshot, x1, y1, x2, y2)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to capture dropdown: {str(e)}")
            self.root.deiconify()
        finally:
            self.is_capturing = False
    
    def open_markup_window(self, screenshot, orig_x, orig_y, orig_x2, orig_y2):
        """Open window for marking up the partial screenshot"""
        # Get screenshot dimensions
        img_width, img_height = screenshot.size
        
        # Set window size based on image size (with padding for controls)
        canvas_width = min(img_width, 800)  # Max width of 800
        canvas_height = img_height
        window_height = canvas_height + 100  # Add space for controls
        
        markup_window = tk.Toplevel(self.root)
        markup_window.title("Partial Capture - Markup")
        markup_window.geometry(f"{canvas_width+10}x{window_height}")
        markup_window.attributes("-topmost", True)
        
        # Current screenshot for display
        current_image = screenshot.copy()
        
        # Convert to PhotoImage for display
        photo = ImageTk.PhotoImage(current_image)
        
        # Create canvas for drawing - size it to the image
        canvas = tk.Canvas(markup_window, bg="white", cursor="crosshair", width=canvas_width, height=canvas_height)
        canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Display image
        img_on_canvas = canvas.create_image(0, 0, image=photo, anchor=tk.NW)
        canvas.image = photo  # Keep a reference
        
        # Markup controls frame
        controls_frame = tk.Frame(markup_window, bg="#f0f0f0", height=50)
        controls_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Tool selection
        tool_var = tk.StringVar(value="rectangle")
        
        tk.Label(controls_frame, text="Tools:", bg="#f0f0f0", font=("Arial", 9)).pack(side=tk.LEFT, padx=5)
        
        tk.Radiobutton(controls_frame, text="Rectangle", variable=tool_var, 
                      value="rectangle", bg="#f0f0f0").pack(side=tk.LEFT, padx=2)
        tk.Radiobutton(controls_frame, text="Circle", variable=tool_var, 
                      value="circle", bg="#f0f0f0").pack(side=tk.LEFT, padx=2)
        tk.Radiobutton(controls_frame, text="Draw", variable=tool_var, 
                      value="draw", bg="#f0f0f0").pack(side=tk.LEFT, padx=2)
        
        # Color selection
        color_var = tk.StringVar(value="red")
        tk.Label(controls_frame, text="Color:", bg="#f0f0f0", font=("Arial", 9)).pack(side=tk.LEFT, padx=(20, 5))
        
        # Color list: red, yellow, and green only
        colors = ["red", "yellow", "green"]
        for color in colors:
            tk.Radiobutton(controls_frame, bg=color, variable=color_var, 
                          value=color, width=2).pack(side=tk.LEFT, padx=2)
        
        # Drawing state
        drawing_state = {'is_drawing': False, 'start_x': 0, 'start_y': 0}
        
        def on_canvas_press(event):
            drawing_state['is_drawing'] = True
            drawing_state['start_x'] = event.x
            drawing_state['start_y'] = event.y
        
        def on_canvas_drag(event):
            if not drawing_state['is_drawing']:
                return
            
            tool = tool_var.get()
            color = color_var.get()
            
            if tool == "draw":
                canvas.create_line(drawing_state['start_x'], drawing_state['start_y'],
                                 event.x, event.y, fill=color, width=2)
                drawing_state['start_x'] = event.x
                drawing_state['start_y'] = event.y
        
        def on_canvas_release(event):
            drawing_state['is_drawing'] = False
            tool = tool_var.get()
            color = color_var.get()
            
            if tool == "circle":
                canvas.create_oval(drawing_state['start_x'], drawing_state['start_y'],
                                 event.x, event.y, outline=color, width=2)
            elif tool == "rectangle":
                canvas.create_rectangle(drawing_state['start_x'], drawing_state['start_y'],
                                      event.x, event.y, outline=color, width=2)
        
        def get_canvas_image():
            """Capture the canvas with all drawings as an image"""
            try:
                # Force canvas to update
                canvas.update()
                time.sleep(0.1)
                
                # Get canvas position and size
                x = canvas.winfo_rootx()
                y = canvas.winfo_rooty()
                w = canvas.winfo_width()
                h = canvas.winfo_height()
                
                # Capture the canvas area from screen
                captured = ImageGrab.grab(bbox=(x, y, x + w, y + h))
                return captured
            except Exception as e:
                print(f"Error capturing canvas: {e}")
                return current_image.copy()
        
        canvas.bind("<Button-1>", on_canvas_press)
        canvas.bind("<B1-Motion>", on_canvas_drag)
        canvas.bind("<ButtonRelease-1>", on_canvas_release)
        
        buttons_frame = tk.Frame(markup_window, bg="#f0f0f0")
        buttons_frame.pack(fill=tk.X, padx=5, pady=5)
        
        def save_markup():
            # Get the marked up image from canvas
            marked_image = get_canvas_image()
            
            # Check if comment checkbox is enabled
            comment = ""
            if self.add_comment_var.get():
                comment = simpledialog.askstring("Add Comment", 
                                               "Enter comment for this partial capture:",
                                               parent=markup_window)
                if comment is None:
                    return  # User cancelled
            
            # Store with markup data
            timestamp = datetime.now()
            self.screenshots.append({
                'image': marked_image,
                'comment': comment if comment else "",
                'timestamp': timestamp,
                'markup_data': {
                    'is_partial': True,
                    'original_coords': (orig_x, orig_y, orig_x2, orig_y2),
                    'tool_used': tool_var.get(),
                    'color_used': color_var.get()
                }
            })
            
            markup_window.destroy()
        
        tk.Button(buttons_frame, text="Save", command=save_markup,
                 bg="#4CAF50", fg="white", font=("Arial", 9)).pack(side=tk.LEFT, padx=5)
        tk.Button(buttons_frame, text="Cancel", command=markup_window.destroy,
                 bg="#f44336", fg="white", font=("Arial", 9)).pack(side=tk.LEFT, padx=5)
    
    def get_comment(self):
        """Get a comment from the user for the screenshot"""
        comment = simpledialog.askstring("Add Comment", 
                                       "Enter a comment for this screenshot (optional):",
                                       parent=self.root)
        return comment if comment is not None else ""
    
    def close_tool(self):
        """Close the tool without saving"""
        if messagebox.askyesno("Close Tool", 
                              "Close SnipIT without saving? Any unsaved screenshots will be lost."):
            self.root.quit()
    
    def show_help(self):
        """Show support and contact information"""
        help_message = """📧 NEED SUPPORT OR HAVE QUERIES?

Please reach me through email:

Work Email:
nanthish.t@gds.ey.com

Personal Email:
nanthishwaran579@gmail.com

I'm here to help with any questions or support you may need!"""
        
        messagebox.showinfo("SnipIT 2.0 - Support", help_message)
        
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
                    doc.add_paragraph("Comment:")
                    doc.add_paragraph(shot['comment'])
                                    
                # Save image to temporary file and add to document
                temp_file = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
                shot["image"].save(temp_file.name, "PNG")
                temp_file.close()
                
                # Add image to document
                doc.add_picture(temp_file.name, width=Inches(5.5))
                
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
            
            # Reset screenshots list
            self.screenshots = []
            
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
