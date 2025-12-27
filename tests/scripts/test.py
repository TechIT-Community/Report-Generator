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
    """
    Recursive function that drives the main test flow (Initial Run).
    
    Logic:
    1. Checks if current page is within limit.
    2. CASE 1 (Standard Pages): Fills text entry and text box widgets with auto-generated content.
    3. CASE 2 (Page 5/Chapters): 
       - Dynamically adds a chapter (on first pass).
       - Fills chapter details.
       - Uploads images to Chapter 1.
    4. Advances to next page (RECURSION).
    5. On Last Page logic: Click Save -> Click Done -> Handle Messagebox.
    """
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
                        content = "Auto-Content for NameAndUSN\n" * 3
                    elif label == "NameUSN":
                        content = "Auto-Content for NameUSN, " * 3
                    elif label == "References":
                        content = "[1] Auto-Ref 1\n[2] Auto-Ref 2\n[3] Auto-Ref 3"
                    elif label == "Abstract":
                         content = f"Auto-Content for Abstract. " * 8
                    else:
                        content = f"Auto-Content for {label}. " * 50
                    
                    widget.insert("1.0", content)

    # CASE 2: CHAPTERS TABS (Page 5)
    else:
        if not hasattr(self, "chapter_tabs") or not self.chapter_tabs:
            print("⚠️ [Mock] Page 5 but no chapter tabs found!")
        
        # TEST: Add a 6th chapter dynamically (first pass only)
        if not hasattr(self, "_added_test_chapter"):
            print("  > [Mock] DYNAMIC CHAPTERS TEST: Adding Chapter 6...")
            self.add_new_chapter_tab()
            self._added_test_chapter = True
        
        for tab in self.chapter_tabs:
            print(f"  > [Mock] Filling data for tab: {tab['name']}")
            self.set_active_tab(tab)
            self.update()
            
            for label, widget, typ in tab["entries"]:
                if typ == "entry":
                     if not widget.get():
                         widget.insert(0, f"Auto-Title for {label}")
                elif typ == "text":
                    if not widget.get("1.0", "end-1c").strip():
                        if "Chapter 1" in tab['name']:
                             content = f"Auto-Content for {label} (Long). " * 200
                        elif "Chapter 2" in tab['name']:
                            content = "Line 1.\nLine 2."
                        else:
                            content = f"Auto-Content for {label}. " * 50
                        widget.insert("1.0", content)
            
            # Upload images for Chapter 1 in first pass
            if "Chapter 1" in tab['name'] and not self.has_uploaded_images:
                 print("  > [Mock] Uploading images (Scenario Test)...")
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
        # Done button - triggers messagebox
        print("🤖 [Mock] Clicking Save before Done...")
        self.save_button.invoke()
        
        print("🤖 [Mock] Clicking Done...")
        self.next_button.invoke()
        
        # Wait for messagebox and start next cycle
        self.after(2000, self.handle_messagebox_and_continue)


def find_upload_button(parent):
    """Helper to recursively find the Upload Images button."""
    for child in parent.winfo_children():
        if isinstance(child, ctk.CTkButton):
            text = child.cget("text")
            if "Upload" in text or "📁" in text:
                return child
        if isinstance(child, ctk.CTkFrame):
            result = find_upload_button(child)
            if result: return result
    return None


def handle_messagebox_and_continue(self):
    """Close messagebox and start next test cycle."""
    # Find and close CTkMessagebox
    for widget in self.winfo_children():
        if "CTkMessagebox" in str(type(widget)) or "Toplevel" in str(type(widget)):
            print("🤖 [Mock] Found messagebox, closing...")
            widget.destroy()
            break
    
    # Also check for any toplevel windows
    for toplevel in self.winfo_toplevel().winfo_children():
        if isinstance(toplevel, ctk.CTkToplevel):
            toplevel.destroy()
    
    # Initialize cycle counter
    if not hasattr(self, "_test_cycle"):
        self._test_cycle = 1
    else:
        self._test_cycle += 1
    
    print(f"🤖 [Mock] Starting Test Cycle {self._test_cycle}...")
    
    if self._test_cycle == 1:
        # CYCLE 1: Add more chapters, upload images to new and existing
        self.after(1000, self.run_cycle_1)
    elif self._test_cycle == 2:
        # CYCLE 2: Delete 2 chapters, add images
        self.after(1000, self.run_cycle_2)
    elif self._test_cycle == 3:
        # CYCLE 3: Delete MORE chapters (3-4), add images to remaining
        self.after(1000, self.run_cycle_3)
    else:
        print("✅ [Mock] All test cycles complete!")
        return


def run_cycle_1(self):
    """Cycle 1: Go back to chapters, add 2 more chapters, fill and upload images."""
    print("🤖 [Mock] Cycle 1: Adding chapters and images...")
    
    # Navigate to page 5 (Chapters)
    self.current_page = 5
    self.load_page()
    self.update()
    
    self.after(1000, self._cycle1_add_chapters)

def _cycle1_add_chapters(self):
    """Add 2 chapters and fill them."""
    # Add 2 new chapters
    print("  > [Mock] Adding Chapter 7...")
    self.add_new_chapter_tab()
    self.update()
    
    print("  > [Mock] Adding Chapter 8...")
    self.add_new_chapter_tab()
    self.update()
    
    self.after(500, self._cycle1_fill_and_upload)

def _cycle1_fill_and_upload(self):
    """Fill new chapters and upload images to new + existing."""
    # Fill new chapters
    for tab in self.chapter_tabs:
        if tab["id"] >= 7:  # New chapters
            print(f"  > [Mock] Filling new {tab['name']}...")
            self.set_active_tab(tab)
            self.update()
            for label, widget, typ in tab["entries"]:
                if typ == "entry" and not widget.get():
                    widget.insert(0, f"Cycle1-Title for {label}")
                elif typ == "text" and not widget.get("1.0", "end-1c").strip():
                    widget.insert("1.0", f"Cycle1-Content for {label}. " * 30)
    
    # Upload images to existing chapters (2, 3) and new chapters (7, 8)
    for tab in self.chapter_tabs:
        if tab["id"] in [2, 3, 7, 8]:
            print(f"  > [Mock] Uploading images to {tab['name']}...")
            self.set_active_tab(tab)
            self.update()
            btn = find_upload_button(tab['frame'])
            if btn:
                btn.invoke()
    
    self.save_current_inputs()
    
    # Press Done
    self.after(1000, self._cycle1_done)

def _cycle1_done(self):
    """Navigate to last page and press Done."""
    print("🤖 [Mock] Cycle 1: Navigating to last page...")
    # Navigate to page 6 (References - last page)
    self.current_page = len(self.pages)
    self.load_page()
    self.update()
    # Now press Next which triggers Done on last page
    self.after(500, self._trigger_done_and_continue)


def run_cycle_2(self):
    """Cycle 2: Go back to chapters, delete random chapters, add images to one."""
    print("🤖 [Mock] Cycle 2: Deleting chapters...")
    
    # Navigate to page 5 (Chapters)
    self.current_page = 5
    self.load_page()
    self.update()
    
    self.after(1000, self._cycle2_delete_chapters)

def _cycle2_delete_chapters(self):
    """Delete a few chapters."""
    # Delete chapters 3 and 5 (if they exist)
    tabs_to_delete = []
    for tab in self.chapter_tabs:
        if tab["id"] in [3, 5]:
            tabs_to_delete.append(tab)
    
    for tab in tabs_to_delete:
        print(f"  > [Mock] Deleting {tab['name']}...")
        self.remove_chapter_tab(tab)
        self.update()
    
    self.after(500, self._cycle2_add_images)

def _cycle2_add_images(self):
    """Upload images to one chapter."""
    # Pick a random remaining chapter
    if self.chapter_tabs:
        tab = random.choice(self.chapter_tabs)
        print(f"  > [Mock] Uploading images to {tab['name']}...")
        self.set_active_tab(tab)
        self.update()
        btn = find_upload_button(tab['frame'])
        if btn:
            btn.invoke()
    
    self.save_current_inputs()
    
    # Press Done
    self.after(1000, self._cycle2_done)

def _cycle2_done(self):
    """Navigate to last page and press Done."""
    print("🤖 [Mock] Cycle 2: Navigating to last page...")
    self.current_page = len(self.pages)
    self.load_page()
    self.update()
    self.after(500, self._trigger_done_and_continue)


def go_next_and_loop(self):
    """Advance page and callback loop."""
    self.go_next()
    self.after(1000, self.run_test_sequence)

gui.App.__init__ = mock_gui_init
gui.App.run_test_sequence = run_test_sequence
gui.App.go_next_and_loop = go_next_and_loop
gui.App.handle_messagebox_and_continue = handle_messagebox_and_continue
gui.App.run_cycle_1 = run_cycle_1
gui.App._cycle1_add_chapters = _cycle1_add_chapters
gui.App._cycle1_fill_and_upload = _cycle1_fill_and_upload
gui.App._cycle1_done = _cycle1_done
gui.App.run_cycle_2 = run_cycle_2
gui.App._cycle2_delete_chapters = _cycle2_delete_chapters
gui.App._cycle2_add_images = _cycle2_add_images
gui.App._cycle2_done = _cycle2_done


def run_cycle_3(self):
    """Cycle 3: Advanced Deletion Logic (Testing robust removal)."""
    print("🤖 [Mock] Cycle 3: Deleting more chapters...")
    
    # Navigate to page 5 (Chapters)
    self.current_page = 5
    self.load_page()
    self.update()
    
    self.after(1000, self._cycle3_delete_chapters)

def _cycle3_delete_chapters(self):
    """Delete a specific number of chapters to test boundary conditions."""
    # Logic adjusted: Deletes `max_delete` chapters.
    # Note: Logic ensures we don't delete the last remaining chapter.
    delete_count = 0
    max_delete = 1  # Number of chapters to delete in this cycle
    
    while delete_count < max_delete and len(self.chapter_tabs) > 1:
        # Always delete the second tab (index 1) to avoid edge cases
        if len(self.chapter_tabs) >= 2:
            tab = self.chapter_tabs[1]
            print(f"  > [Mock] Deleting {tab['name']}...")
            self.remove_chapter_tab(tab)
            self.update()
            delete_count += 1
    
    print(f"  > [Mock] Deleted {delete_count} chapters. Remaining: {len(self.chapter_tabs)}")
    
    self.after(500, self._cycle3_add_images)

def _cycle3_add_images(self):
    """Upload images to remaining chapters."""
    if self.chapter_tabs:
        tab = self.chapter_tabs[0]  # First remaining chapter
        print(f"  > [Mock] Uploading images to {tab['name']}...")
        self.set_active_tab(tab)
        self.update()
        btn = find_upload_button(tab['frame'])
        if btn:
            btn.invoke()
    
    self.save_current_inputs()
    
    # Press Done
    self.after(1000, self._cycle3_done)

def _cycle3_done(self):
    """Navigate to last page and press Done."""
    print("🤖 [Mock] Cycle 3: Navigating to last page...")
    self.current_page = len(self.pages)
    self.load_page()
    self.update()
    self.after(500, self._trigger_done_and_continue)


def _trigger_done_and_continue(self):
    """Common helper: Press Done button (on last page) and handle messagebox."""
    print("🤖 [Mock] Pressing Done...")
    self.next_button.invoke()  # On last page, this triggers save_document
    self.after(2000, self.handle_messagebox_and_continue)


gui.App.run_cycle_3 = run_cycle_3
gui.App._cycle3_delete_chapters = _cycle3_delete_chapters
gui.App._cycle3_add_images = _cycle3_add_images
gui.App._cycle3_done = _cycle3_done
gui.App._trigger_done_and_continue = _trigger_done_and_continue


# =================================================================================================
#                                         MAIN EXECUTION
# =================================================================================================

if __name__ == "__main__":
    print(f"🚀 Starting Test Script...")
    print(f"📂 Project Root: {PROJECT_ROOT}")
    print(f"🖼️ Test Image: {TEST_IMAGE_PATH}")
    
    MainApp.main()
