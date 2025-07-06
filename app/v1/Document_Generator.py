"""
Backend Document Generator for Reports
This module handles the creation and management of Word documents for report generation.
It uses the `win32com` library to interact with Microsoft Word, allowing for dynamic content
insertion, formatting, and saving of documents.
It also provides utility functions for setting up document properties, inserting static content,
and replacing bookmarks with user-provided data.
It is designed to be used in conjunction with a GUI application that collects user inputs
and passes them to this module for document generation.    
"""

# =================================================================================

# Imports
import win32com.client as win32 # For interacting with Microsoft Word
from win32com.client import constants as c # Constants for Word operations
from pathlib import Path # For path management
import win32gui # For GUI window management
import win32con # For window constants
import time # for pauses
import ctypes # For getting screen dimensions

# =================================================================================

# Globals
word = win32.gencache.EnsureDispatch("Word.Application") # Launch Word and Ensure that its running
word.Visible = True # Show Word window
DOC_PATH = Path.cwd() / "app" / "v1" / "reports" / "template.docx" # Save location
doc = word.Documents.Add() # Create a new document

# Setup Word window
hwnd = win32gui.FindWindow("OpusApp", None) # Find the Word window
win32gui.ShowWindow(hwnd, win32con.SW_RESTORE) # Restore the window if minimized
win32gui.SetForegroundWindow(hwnd) # Bring Word to the foreground

# =================================================================================

# Helper Functions
cm_to_pt = lambda cm: cm * 28.3464566929133858 # For point system in word (1 cm = 28.346 pt)

# Simulates n Backspace key presses
def backspace(n=1):
    sel = word.Selection
    if sel.Start >= n:
        backspace_range = doc.Range(sel.Start - n, sel.Start)
        backspace_range.Delete()

# =================================================================================

# Set margins
doc.PageSetup.TopMargin = cm_to_pt(1.7)
doc.PageSetup.BottomMargin = cm_to_pt(1.7)
doc.PageSetup.LeftMargin = cm_to_pt(2.1)
doc.PageSetup.RightMargin = cm_to_pt(1.7)

# Delete any default text
doc.Content.Delete()

# Global cursor
cursor = doc.Range(0, 0)
cursor.Collapse(c.wdCollapseEnd)

# =================================================================================

def position_windows():
    """
    Positions the Word window and the GUI application side by side for better usability.
    This function calculates the screen dimensions and sets the Word window to occupy
    the left half of the screen, adjusting its size and position accordingly.
    It also sets the zoom level of the Word document to 110% and scrolls to the middle.
    """
    screen_width = ctypes.windll.user32.GetSystemMetrics(0) #1920
    screen_height = ctypes.windll.user32.GetSystemMetrics(1) #1080

    half_width = screen_width // 2
    height = int(screen_height * 0.99)

    left = int(max(0, half_width - 0.107 * screen_width))
    width = int((half_width + 0.11 * screen_width))

    hwnd_word = win32gui.FindWindow("OpusApp", None) # Find the Word window
    if hwnd_word:
        win32gui.ShowWindow(hwnd_word, win32con.SW_RESTORE) # Restore the window if minimized
        # Set position and size
        win32gui.SetWindowPos( 
            hwnd_word, None,
            left, 0,
            width, height,
            win32con.SWP_NOZORDER
        ) 

    word.ActiveWindow.View.Zoom.Percentage = 110 # Change zoom level
    window = word.ActiveWindow # Get the active window
    window.ScrollIntoView(doc.Range(0, doc.Content.End // 2), True) # Scroll to the middle of the document

# ---------------------------------------------------------------------------------

def insert_static_content():
    """
    Inserts static content into the Word document and adds placeholders for dynamic content..
    This function makes sure to set the font, size and alignment appropriately for the heading,
    sub-heading, and content before insertion, even for placeholders.
    """
    position_windows()  # Call to arrange Word window properly
# _________________________________________________________________________________

    global cursor
    cursor.Select()
    
    word.Selection.Font.Name = "Times New Roman"                            # Font Name
    word.Selection.Font.Size = 15                                           # Font Size
    word.Selection.Font.Bold = True                                         # Bold        
    word.Selection.Font.Italic = False                                      # Italic  
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter     # Alignment
    word.Selection.ParagraphFormat.LineSpacingRule = c.wdLineSpaceSingle    # Line Spacing

    word.Selection.TypeText(
        "VISVESVARAYA TECHNOLOGICAL UNIVERSITY\n"
        "“Jnana Sangama”, Belagavi – 590 018"
    )
    word.Selection.TypeParagraph()
    time.sleep(0.1)
# _________________________________________________________________________________

    cursor = word.Selection.Range # Get the current selection range
    cursor.Collapse(c.wdCollapseEnd) # Move cursor to the end
    cursor.InsertAfter("\n")
    cursor.Collapse(c.wdCollapseStart) # Move cursor to the start
    
    image_path = str(Path.cwd() / "app" / "v1" / "assets" / "VTU_Logo.png")
    cursor.InsertParagraphAfter() # Insert a paragraph break
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter

    inline_shape = doc.InlineShapes.AddPicture(image_path, False, True, cursor) # Insert the image
    inline_shape.LockAspectRatio = True # Lock aspect ratio
    inline_shape.Width = cm_to_pt(4) # Set width to 4 cm

    cursor = inline_shape.Range.Duplicate # Duplicate the range of the inserted image
    cursor.Collapse(c.wdCollapseEnd) 
    cursor.InsertParagraphAfter() # Insert a paragraph break after the image
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
# _________________________________________________________________________________

    word.Selection.Font.Size = 10
    word.Selection.TypeText("A MINI PROJECT\non\n")
    time.sleep(0.1)
# _________________________________________________________________________________

    word.Selection.TypeText("___\n")
    title_range = word.Selection.Range.Duplicate # start range to bookmark for Project Title
    title_range.MoveStart(Unit=c.wdCharacter, Count=-4) # Length of bookmark (from end, backwards)
    doc.Bookmarks.Add("ProjectTitle", title_range) # Bookmark for Project Title (Placeholder)
    time.sleep(0.1)
# _________________________________________________________________________________

    word.Selection.Font.Size = 10
    word.Selection.Font.Bold = False
    word.Selection.Font.Italic = True
    word.Selection.TypeText("Submitted in partial fulfillment of the requirements for the award of degree")
    word.Selection.TypeParagraph()
    time.sleep(0.1)
# _________________________________________________________________________________


# ---------------------------------------------------------------------------------

def replace_bookmarks(data_dict: dict):
    """
    Replaces bookmarks in the Word document with values from a dictionary.
    This function iterates through the provided dictionary and checks if each key exists as a bookmark in the document.
    If a bookmark exists, it replaces the text of that bookmark with the corresponding value from the dictionary.

    :param data_dict: A dictionary where keys are bookmark names and values are the text to replace them with.
    :type data_dict: dict
    """
    for key, value in data_dict.items():
        if doc.Bookmarks.Exists(key):
            bm_range = doc.Bookmarks(key).Range # range of bookmark
            bm_start = bm_range.Start # start position of bookmark
            bm_range.Text = value + "\n" # Replace bookmark text with value
            new_range = doc.Range(bm_start, bm_start + len(value) + 1) # create a new range for the bookmark
            doc.Bookmarks.Add(key, new_range) # Re-add the bookmark with the new range
    cursor = doc.Range()
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()

# =================================================================================

def save_document():
    """
    Saves the current Word document to the specified path.
    """
    doc.SaveAs(str(DOC_PATH), FileFormat=c.wdFormatDocumentDefault)
    print("✅ Saved:", DOC_PATH.resolve())
    
# =================================================================================