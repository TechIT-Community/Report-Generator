
import sys
import os
import random
import time
import threading
from pathlib import Path
import tkinter.filedialog
import customtkinter as ctk

# Add the project root to sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.append(str(PROJECT_ROOT))
sys.path.append(str(PROJECT_ROOT / "app"))

# Mock assets
PATHS = [PROJECT_ROOT / "screenshots" / "test1.png",
 PROJECT_ROOT / "screenshots" / "test2.png",
 PROJECT_ROOT / "screenshots" / "test3.png"]

TEST_IMAGE_PATH = random.choice(PATHS)

# =================================================================================
# Monkey Patching
# =================================================================================

# 1. Patch file dialog to always return our test image
original_askopenfilenames = tkinter.filedialog.askopenfilenames

def mock_askopenfilenames(*args, **kwargs):
    print("🤖 [Mock] File dialog opened. Auto-selecting 5 test images.")
    return [str(random.choice(PATHS)) for _ in range(5)]

tkinter.filedialog.askopenfilenames = mock_askopenfilenames

# 2. Patch StartScreen to auto-select and launch
import app.Main as MainApp

original_start_init = MainApp.StartScreen.__init__

def mock_start_init(self, *args, **kwargs):
    original_start_init(self, *args, **kwargs)
    print("🤖 [Mock] StartScreen initialized. selecting defaults...")
    
    self.after(1000, lambda: self._auto_start())

def _auto_start(self):
    print("🤖 [Mock] Auto-starting application...")
    self.college_var.set("BNMIT")
    self.dept_var.set("COMPUTER SCIENCE AND ENGINEERING")
    
    # Add a print here because the next step (import gui) takes time
    print("⏳ [Mock] Launching GUI... (This connects to Word, so it may pause here for a few seconds)")
    self.start_app()

MainApp.StartScreen.__init__ = mock_start_init
MainApp.StartScreen._auto_start = _auto_start


# 3. Patch Main GUI to auto-fill
import gui

original_gui_init = gui.App.__init__

def mock_gui_init(self, user_inputs):
    original_gui_init(self, user_inputs)
    print("🤖 [Mock] Main GUI initialized. Starting auto-pilot...")
    self.has_uploaded_images = False # Flag to track if we've tested uploads
    self.after(2000, self.run_test_sequence)

def run_test_sequence(self):
    if self.current_page > len(self.pages):
        print("🤖 [Mock] Reached end of pages. Test complete.")
        return

    print(f"🤖 [Mock] Processing Page {self.current_page}...")
    
    # Fill inputs
    for label, widget, typ in self.entries:
        if typ == "entry":
             current_text = widget.get()
             if not current_text:
                 widget.insert(0, f"Auto-Text for {label}")
        elif typ == "text":
            current_text = widget.get("1.0", "end-1c")
            if not current_text.strip():
                if label == "NameAndUSN":
                    # Specific format: 3 lines
                    content = "Auto-Content for NameAndUSN\n" * 3
                elif label == "NameUSN":
                    # Specific format: 3 times in one line (mostly)
                    content = "Auto-Content for NameUSN, " * 3
                elif label == "References":
                    # Explicit references content
                    content = "[1] Auto-Ref 1\n[2] Auto-Ref 2\n[3] Auto-Ref 3"
                else:
                    # Variable length: Make Chapter 2 (Page 5, index 4 in 0-based list but here check page num)
                    # Page 1-3 info, Page 4=Ch1, Page 5=Ch2
                    if self.current_page == 4:
                         # Abstract: Normal/Short
                         content = f"Auto-Content for Abstract. " * 8
                    elif self.current_page == 5:
                        # Chapter 1: Long
                         content = f"Auto-Content for {label} (Long Version). " * 200
                    elif self.current_page == 6:
                        # Chapter 2: Short (2 lines)
                        content = "Line 1 of Chapter 2.\nLine 2 of Chapter 2."
                    elif self.current_page == 7:
                        # Chapter 3: Medium (8 lines)
                        content = "Line 1 of Chapter 3.\nLine 2.\nLine 3.\nLine 4.\nLine 5.\nLine 6.\nLine 7.\nLine 8 of Chapter 3."
                    else:
                        # Chapter 4, 5: Normal
                        content = f"Auto-Content for {label}. " * 50
                
                widget.insert("1.0", content)
    
    # Upload images logic
    # Page 1-3 are info, 4-8 are chapters 1-5
    # So valid image pages are page indices 4,5,6,7,8 (1-indexed)
    
    is_chapter_page = 4 <= self.current_page <= 8
    
    if is_chapter_page:
        # ALWAYS upload for testing smart placement
        self.has_uploaded_images = False # Reset flag for each page
        if True: # Always runs
            print("🤖 [Mock] Triggering Image Upload for this chapter...")
            
            # Find the upload button. It's buried in the widget hierarchy.
            # In gui.py: self.input_frame -> image_upload_frame -> upload_btn
            
            # Helper to recursively find button with text "Upload Images"
            def find_upload_button(parent):
                for child in parent.winfo_children():
                    if isinstance(child, ctk.CTkButton) and child.cget("text") == "Upload Images":
                        return child
                    if isinstance(child, ctk.CTkFrame):
                        result = find_upload_button(child)
                        if result: return result
                return None

            btn = find_upload_button(self.input_frame)
            if btn:
                btn.invoke()
                self.has_uploaded_images = True
            else:
                print("⚠️ [Mock] Could not find upload button!")

    # Move Next
    if self.current_page < len(self.pages):
        self.after(1000, self.go_next_and_loop)
    else:
        # Done button
        print("🤖 [Mock] Clicking Save before Done...")
        self.save_button.invoke()
        
        print("🤖 [Mock] Clicking Done...")
        self.after(1000, self.next_button.invoke)

def go_next_and_loop(self):
    self.go_next()
    self.after(1000, self.run_test_sequence)

gui.App.__init__ = mock_gui_init
gui.App.run_test_sequence = run_test_sequence
gui.App.go_next_and_loop = go_next_and_loop


# =================================================================================
# Main Execution
# =================================================================================

if __name__ == "__main__":
    print(f"🚀 Starting Test Script...")
    print(f"📂 Project Root: {PROJECT_ROOT}")
    print(f"🖼️ Test Image: {TEST_IMAGE_PATH}")
    
    MainApp.main()
