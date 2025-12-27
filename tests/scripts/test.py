"""
Automated Test Script for Report Generator.
Mocks user interactions (clicks, typing, file selection) to verify the end-to-end flow.

Key Features:
- Monkey patches `tkinter.filedialog` to auto-select images.
- Monkey patches `StartScreen` to auto-fill college/dept and proceed.
- Monkey patches `MainApp` logic to simulate typing into all fields (Long/Short text).
- Verifies image placement logic by forcibly uploading images in every chapter.

Usage:
    python tests/scripts/test.py
"""

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

# =================================================================================================
#                                       MONKEY PATCHES
# =================================================================================================

# -------------------------- Patch 1: File Dialog --------------------------

# Patch file dialog to always return our test image
original_askopenfilenames = tkinter.filedialog.askopenfilenames

def mock_askopenfilenames(*args, **kwargs):
    """Intercepts file dialog to return random test images without user interaction."""
    print("🤖 [Mock] File dialog opened. Auto-selecting 5 test images.")
    # Return 5 random images from the available set
    return [str(random.choice(PATHS)) for _ in range(5)]

tkinter.filedialog.askopenfilenames = mock_askopenfilenames


# -------------------------- Patch 2: Start Screen --------------------------

import app.frontend.main as MainApp # Updated import path

original_start_init = MainApp.StartScreen.__init__

def mock_start_init(self, *args, **kwargs):
    """Intercepts StartScreen init to trigger auto-start sequence."""
    original_start_init(self, *args, **kwargs)
    print("🤖 [Mock] StartScreen initialized. selecting defaults...")
    self.after(1000, lambda: self._auto_start())

def _auto_start(self):
    """Simulates selecting dropdowns and clicking Start."""
    print("🤖 [Mock] Auto-starting application...")
    self.college_var.set("BNMIT")
    self.dept_var.set("COMPUTER SCIENCE AND ENGINEERING")
    
    # Add a print here because the next step (import gui) takes time
    print("⏳ [Mock] Launching GUI... (This connects to Word, so it may pause here for a few seconds)")
    self.start_app()

MainApp.StartScreen.__init__ = mock_start_init
MainApp.StartScreen._auto_start = _auto_start


# -------------------------- Patch 3: Main GUI Logic --------------------------

import app.frontend.gui as gui # Updated import path

original_gui_init = gui.App.__init__

def mock_gui_init(self, user_inputs):
    """Intercepts Main GUI init to trigger the test sequence."""
    original_gui_init(self, user_inputs)
    print("🤖 [Mock] Main GUI initialized. Starting auto-pilot...")
    self.has_uploaded_images = False # Flag to track if we've tested uploads
    self.after(2000, self.run_test_sequence)

def run_test_sequence(self):
    """Recursive function to fill page inputs, upload images, and click Next."""
    if self.current_page > len(self.pages):
        print("🤖 [Mock] Reached end of pages. Test complete.")
        return

    print(f"🤖 [Mock] Processing Page {self.current_page}...")
    
    # CASE 1: STANDARD PAGES (Not 5)
    if self.current_page != 5:
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
                    elif label == "Abstract":
                         # Abstract: Normal/Short
                         content = f"Auto-Content for Abstract. " * 8
                    else:
                        content = f"Auto-Content for {label}. " * 50
                    
                    widget.insert("1.0", content)

    # CASE 2: CHAPTERS TABS (Page 5)
    else:
        # Iterate through all tabs
        # Need to access self.chapter_tabs (from gui.py logic)
        
        # NOTE: self.chapter_tabs is available because we patched methods onto the instance
        if not hasattr(self, "chapter_tabs") or not self.chapter_tabs:
            print("⚠️ [Mock] Page 5 but no chapter tabs found!")
        
        # TEST: Add a 6th chapter dynamically
        if not hasattr(self, "_added_test_chapter"):
            print("  > [Mock] DYNAMIC CHAPTERS TEST: Adding Chapter 6...")
            self.add_new_chapter_tab()
            self._added_test_chapter = True

        
        for tab in self.chapter_tabs:
            print(f"  > [Mock] Filling data for tab: {tab['name']}")
            
            # Switch to this tab to ensure widgets are reliable (though they exist in memory regardless)
            self.set_active_tab(tab) 
            self.update() # Force UI refresh
            
            # Fill inputs in THIS tab
            # tab['entries'] stores (label, widget, type)
            for label, widget, typ in tab["entries"]:
                if typ == "entry":
                     if not widget.get():
                         widget.insert(0, f"Auto-Title for {label}")
                elif typ == "text":
                    if not widget.get("1.0", "end-1c").strip():
                        # Vary content length based on chapter
                        if "Chapter 1" in tab['name']:
                             content = f"Auto-Content for {label} (Long). " * 200
                        elif "Chapter 2" in tab['name']:
                            content = "Line 1.\nLine 2."
                        else:
                            content = f"Auto-Content for {label}. " * 50
                        widget.insert("1.0", content)
            
            # Upload Images Logic for THIS tab
            if "Chapter 1" in tab['name']:
                 print("  > [Mock] Uploading images (Scenario Test)...")
                 # Reuse the upload verification logic
                 
                 # Helper to recursively find button with text "Upload Images"
                 # Since we are in a tab, search inside tab['frame']
                 def find_upload_button(parent):
                    for child in parent.winfo_children():
                        if isinstance(child, ctk.CTkButton) and child.cget("text") == "Upload Images":
                            return child
                        if isinstance(child, ctk.CTkFrame):
                            result = find_upload_button(child)
                            if result: return result
                    return None
                 
                 btn = find_upload_button(tab['frame'])
                 if btn:
                     btn.invoke()
                     self.has_uploaded_images = True
                 else:
                     print(f"⚠️ [Mock] Could not find upload button in {tab['name']}!")

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
    """advance page and callback loop."""
    self.go_next()
    self.after(1000, self.run_test_sequence)

gui.App.__init__ = mock_gui_init
gui.App.run_test_sequence = run_test_sequence
gui.App.go_next_and_loop = go_next_and_loop


# =================================================================================================
#                                         MAIN EXECUTION
# =================================================================================================

if __name__ == "__main__":
    print(f"🚀 Starting Test Script...")
    print(f"📂 Project Root: {PROJECT_ROOT}")
    print(f"🖼️ Test Image: {TEST_IMAGE_PATH}")
    
    MainApp.main()
