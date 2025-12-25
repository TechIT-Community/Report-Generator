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
# =================================================================================

# Imports
import win32com.client as win32 # For interacting with Microsoft Word
from win32com.client import constants as c # Constants for Word operations
from pathlib import Path # For path management
import win32gui # For GUI window management
import win32con # For window constants
import time # for pauses
import ctypes # For getting screen dimensions
from CTkMessagebox import CTkMessagebox
import re
from PIL import Image

# ================================================================================= 
# =================================================================================

BASE_DIR = Path(__file__).resolve().parent  # Base directory of the application
ASSET_DIR = BASE_DIR / "assets"  # Directory for assets 

# Globals
word = win32.gencache.EnsureDispatch("Word.Application") # Launch Word and Ensure that its running
word.Visible = True # Show Word window
DOC_PATH = BASE_DIR / "reports" / "template.docx" # Save location
doc = word.Documents.Add() # Create a new document

# Setup Word window
hwnd = win32gui.FindWindow("OpusApp", None) # Find the Word window
win32gui.ShowWindow(hwnd, win32con.SW_RESTORE) # Restore the window if minimized
win32gui.SetForegroundWindow(hwnd) # Bring Word to the foreground

# ================================================================================= 
# =================================================================================

# Helper Functions
cm_to_pt = lambda cm: cm * 28.3464566929133858 # For point system in word (1 cm = 28.346 pt)

def set_format(font_name=None, size=None, bold=None, italic=None, align=None, underline=None):
    """
    Sets the formatting for the current selection in Word. Only applies provided values.
    """
    if font_name is not None: word.Selection.Font.Name = font_name
    if size is not None: word.Selection.Font.Size = size
    if bold is not None: word.Selection.Font.Bold = bold
    if italic is not None: word.Selection.Font.Italic = italic
    if align is not None: word.Selection.ParagraphFormat.Alignment = align
    if underline is not None: word.Selection.Font.Underline = underline

def add_bookmark(name, placeholder="___", add_newline=False):
    """
    Types a placeholder, wraps it in a bookmark, and optionally adds a newline or space.
    """
    word.Selection.TypeText(placeholder)
    bm_range = word.Selection.Range.Duplicate
    bm_start = bm_range.Start - len(placeholder)
    bm_range = doc.Range(bm_start, bm_start + len(placeholder))
    doc.Bookmarks.Add(name, bm_range)
    if add_newline:
        word.Selection.TypeParagraph()

# Simulates n Backspace key presses
def backspace(n=1):
    sel = word.Selection
    if sel.Start >= n:
        backspace_range = doc.Range(sel.Start - n, sel.Start)
        backspace_range.Delete()
        

def insert_table(data: list[list[str]], bold_cells: list[tuple[int, int]] = None, align = c.wdAlignParagraphCenter, before = 0, after = 8, transparent = False):
    """
    Inserts a table into the Word document with data oriented as-is (row-wise).
    
    Args:
        data (list[list[str]]): Row-wise list of lists. Each sublist is a row.
        bold_cells (list[tuple[int, int]]): List of (row_index, col_index) tuples to make bold.
        align (int): Paragraph alignment.
        before (int): Space before paragraph.
        after (int): Space after paragraph.
        transparent (bool): Whether borders are invisible.
    """
    global cursor

    if not data or not any(data):
        return

    bold_cells = bold_cells or []

    rows = len(data)
    cols = max(len(row) for row in data)

    # Normalize data: pad missing cells with empty strings
    normalized_data = []
    for row in data:
        normalized_row = []
        for j in range(cols):
            val = row[j] if j < len(row) else ""
            clean_val = "" if not val or str(val).strip() == "" else str(val)
            normalized_row.append(clean_val)
        normalized_data.append(normalized_row)

    # Insert at end
    cursor = doc.Range()
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()

    table = doc.Tables.Add(cursor, NumRows=rows, NumColumns=cols)
    table.Range.Style = "Table Grid"

    # Global table formatting
    table.Range.Font.Name = "Times New Roman"
    table.Range.Font.Size = 12
    table.Range.ParagraphFormat.Alignment = align
    table.Range.ParagraphFormat.LineSpacingRule = c.wdLineSpaceSingle
    table.Range.ParagraphFormat.SpaceBefore = before
    table.Range.ParagraphFormat.SpaceAfter = after

    # Fill content and apply bold
    for i, row in enumerate(normalized_data):
        for j, cell_val in enumerate(row):
            cell = table.Cell(i + 1, j + 1)
            cell.Range.Text = cell_val
            if (i, j) in bold_cells:
                cell.Range.Font.Bold = True

    # Apply borders
    color = c.wdColorWhite if transparent else c.wdColorBlack
    for border_id in [
        c.wdBorderTop, c.wdBorderBottom, c.wdBorderLeft, c.wdBorderRight,
        c.wdBorderHorizontal, c.wdBorderVertical
    ]:
        border = table.Borders(border_id)
        border.LineStyle = c.wdLineStyleSingle
        border.Color = color

    # Move cursor after table
    cursor = table.Range.Duplicate
    cursor.Collapse(c.wdCollapseEnd)
    cursor.InsertParagraphAfter()
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()


# ================================================================================= 
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

# Enforce global font setting for Normal style and defaults
try:
    doc.Styles(c.wdStyleNormal).Font.Name = "Times New Roman"
    doc.Content.Font.Name = "Times New Roman"
    # Also ensure Default Paragraph Font is checked if possible, but doc.Content usually covers it.
except:
    pass


# ================================================================================= 
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
# _________________________________________________________________________________
    
    set_format(size=15, bold=True, align=c.wdAlignParagraphCenter, underline=c.wdUnderlineNone)


    word.Selection.TypeText(
        "VISVESVARAYA TECHNOLOGICAL UNIVERSITY\n"
        "“Jnana Sangama”, Belagavi – 590 018"
    )
    word.Selection.TypeParagraph()
    # time.sleep(0.1)
# _________________________________________________________________________________

    cursor = word.Selection.Range # Get the current selection range
    cursor.Collapse(c.wdCollapseEnd) # Move cursor to the end
    word.Selection.TypeParagraph() 
    cursor.Collapse(c.wdCollapseStart) # Move cursor to the start
    
    image_path = str(BASE_DIR / "assets" / "VTU_Logo.png")
#    cursor.InsertParagraphAfter() # Insert a paragraph break
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter

    inline_shape = doc.InlineShapes.AddPicture(image_path, False, True, cursor) # Insert the image
    inline_shape.LockAspectRatio = True # Lock aspect ratio
    inline_shape.Width = cm_to_pt(4) # Set width to 4 cm

    cursor = inline_shape.Range.Duplicate # Duplicate the range of the inserted image
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    # cursor.InsertParagraphAfter() # Insert a paragraph break after the image
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
    # time.sleep(0.1)
# _________________________________________________________________________________

    word.Selection.Font.Size = 11
    word.Selection.TypeText("A MINI PROJECT\vOn")
    word.Selection.TypeParagraph()
    # time.sleep(0.1)
# _________________________________________________________________________________
    
    set_format(size=15, bold=True, align=c.wdAlignParagraphCenter)
    add_bookmark("ProjectTitle", "___\n")
    # time.sleep(0.1)
# _________________________________________________________________________________

    set_format(size=11, bold=False, italic=True, align=c.wdAlignParagraphCenter)
    word.Selection.TypeText("Submitted in partial fulfilment of the requirements for the award of degree")
    word.Selection.TypeParagraph()
    # time.sleep(0.1)
# _________________________________________________________________________________

    set_format(size=11, bold=False, italic=False, align=c.wdAlignParagraphCenter)
    word.Selection.TypeText("Bachelor of Engineering\vIn\v")
    # time.sleep(0.1)

    word.Selection.Font.Bold = True
    add_bookmark("Department", "___")
    word.Selection.TypeParagraph()    

    word.Selection.Font.Bold = False
    word.Selection.TypeText("Submitted by")
    word.Selection.TypeParagraph()    
    # time.sleep(0.1)
# _________________________________________________________________________________

    word.Selection.Font.Bold = True
    add_bookmark("NameAndUSN", "___\n")
    # time.sleep(0.1)
# _________________________________________________________________________________

    word.Selection.Font.Bold = False
    word.Selection.TypeText("Under the guidance of\v")    
    # time.sleep(0.1)
# _________________________________________________________________________________
    
    word.Selection.Font.Bold = True
    add_bookmark("GuideName", "___\n")
    # word.Selection.TypeParagraph() # Removed to prevent double newline
 
    # time.sleep(0.1)
# _________________________________________________________________________________

    word.Selection.Font.Bold = False
    add_bookmark("Designation", "___\n")
    # time.sleep(0.1)
# _________________________________________________________________________________

    cursor = word.Selection.Range 
    cursor.Collapse(c.wdCollapseEnd) 
    word.Selection.TypeParagraph() 
    cursor.Collapse(c.wdCollapseStart)
    
    image_path = str(BASE_DIR / "assets" / "BNMIT_Logo.png")
#    cursor.InsertParagraphAfter() 
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter

    inline_shape = doc.InlineShapes.AddPicture(image_path, False, True, cursor) 
    inline_shape.LockAspectRatio = True 
    inline_shape.Width = cm_to_pt(5) 

    cursor = inline_shape.Range.Duplicate 
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    # cursor.InsertParagraphAfter()
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
    # time.sleep(0.1)
# _________________________________________________________________________________

    word.Selection.Font.Bold = True
    add_bookmark("Department_2", "___\n")
    doc.Bookmarks("Department_2").Range.Case = c.wdUpperCase 
    # time.sleep(0.1)
# _________________________________________________________________________________

    cursor = word.Selection.Range 
    cursor.Collapse(c.wdCollapseEnd) 
    
    image_path = str(BASE_DIR / "assets" / "BNMIT_Text.png")
#    cursor.InsertParagraphAfter() 
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter

    inline_shape = doc.InlineShapes.AddPicture(image_path, False, True, cursor) 
    inline_shape.LockAspectRatio = True 
    inline_shape.Width = cm_to_pt(15) 

    cursor = inline_shape.Range.Duplicate 
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    # cursor.InsertParagraphAfter()
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
    # time.sleep(0.1)
# _________________________________________________________________________________
# _________________________________________________________________________________

    cursor.Collapse(c.wdCollapseEnd)
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1) 
    
    cursor.InsertBreak(c.wdPageBreak)
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
    # time.sleep(0.1) 
# _________________________________________________________________________________
# _________________________________________________________________________________

    cursor = word.Selection.Range 
    cursor.Collapse(c.wdCollapseEnd)
    
    image_path = str(BASE_DIR / "assets" / "BNMIT_Text.png")
#    cursor.InsertParagraphAfter() 
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter

    inline_shape = doc.InlineShapes.AddPicture(image_path, False, True, cursor) 
    inline_shape.LockAspectRatio = True 
    inline_shape.Width = cm_to_pt(15) 

    cursor = inline_shape.Range.Duplicate 
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    cursor.InsertParagraphAfter()
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
    # time.sleep(0.1)
# _________________________________________________________________________________

    placeholder = "___\n"
    word.Selection.TypeText(placeholder)
    bm_range = word.Selection.Range.Duplicate
    bm_start = bm_range.Start - len(placeholder)
    bm_range = doc.Range(bm_start, bm_start + len(placeholder))
    doc.Bookmarks.Add("Department_3", bm_range)
    bm_range.Case = c.wdUpperCase 
    # time.sleep(0.1)
# _________________________________________________________________________________

    cursor = word.Selection.Range 
    cursor.Collapse(c.wdCollapseEnd) 
    word.Selection.TypeParagraph()
    cursor.Collapse(c.wdCollapseStart)
    
    image_path = str(BASE_DIR / "assets" / "BNMIT_Logo.png")
    cursor.InsertParagraphAfter() 
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter

    inline_shape = doc.InlineShapes.AddPicture(image_path, False, True, cursor) 
    inline_shape.LockAspectRatio = True 
    inline_shape.Width = cm_to_pt(5) 

    cursor = inline_shape.Range.Duplicate 
    cursor.Collapse(c.wdCollapseEnd) 
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
    # time.sleep(0.1)
# _________________________________________________________________________________

    word.Selection.Font.Name = "Calibri"                           
    word.Selection.Font.Size = 15                                          
    word.Selection.Font.Bold = True                                                
    word.Selection.Font.Italic = False                                       
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter     
    word.Selection.ParagraphFormat.LineSpacingRule = c.wdLineSpace1pt5    
    word.Selection.Font.Underline = c.wdUnderlineSingle

    word.Selection.TypeText("CERTIFICATE")
    word.Selection.TypeParagraph()
    # time.sleep(0.1)
# _________________________________________________________________________________

    word.Selection.Font.Name = "Times New Roman"                            
    word.Selection.Font.Size = 12                                          
    word.Selection.Font.Bold = False                                                
    word.Selection.Font.Italic = False                                       
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphJustify     
    word.Selection.ParagraphFormat.LineSpacingRule = c.wdLineSpace1pt5    
    word.Selection.Font.Underline = c.wdUnderlineNone

    word.Selection.TypeText("This is to certify that the Mini project work entitled ")
    set_format(underline=c.wdUnderlineNone)
    # time.sleep(0.1)
# _________________________________________________________________________________
    
    set_format(bold=True)
    add_bookmark("ProjectTitle_2", "___")
    
    set_format(bold=False)
    word.Selection.TypeText(" is a bonafide work carried out by ")

    set_format(bold=True)
    add_bookmark("NameAndUSN_2", "___\n")
    
    set_format(bold=False)
    word.Selection.TypeText(" in partial fulfilment for the award of degree of ")

    set_format(bold=True)
    word.Selection.TypeText("Bachelor of Engineering")
    set_format(bold=False)
    word.Selection.TypeText(" in ")
    set_format(bold=True)
    add_bookmark("Department_4", "___") # Changed from Department_3 to match original logic if distinct
    
    set_format(bold=False)
    word.Selection.TypeText(" of the ")
    set_format(bold=True)
    word.Selection.TypeText("Visvesvaraya Technological University, Belagavi")
    set_format(bold=False)
    word.Selection.TypeText(" during the year ")
    set_format(bold=True)
    add_bookmark("Year", "___")
    
    set_format(bold=False)
    word.Selection.TypeText(". It is certified that all corrections/suggestions indicated for Internal Assessment have been incorporated in the report deposited in the departmental library. The project report has been approved as it satisfies the academic requirements in respect of Project work prescribed for the said Degree.")
    # time.sleep(0.1)
# _________________________________________________________________________________

    data = [
        ["___",     "___", "Dr. S Y Kulkarni"],
        ["___,",       "Professor and HOD,", "Additional Director"],
        ["___,",     "___,",      "and Principal,"],
        ["BNMIT, Bengaluru", "BNMIT, Bengaluru",   "BNMIT, Bengaluru"]
    ]
    
    bold_cells = [(0, 0), (0, 1), (0, 2)]

    rows = len(data)
    cols = max(len(row) for row in data)

    cursor.Collapse(c.wdCollapseEnd)
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1) 
    cursor.Select()

    table = doc.Tables.Add(cursor, NumRows=rows, NumColumns=cols)
    table.Range.Style = "Table Grid"
    
    table.Range.Font.Name = "Times New Roman"
    table.Range.Font.Size = 12
    table.Range.ParagraphFormat.Alignment = c.wdAlignParagraphCenter
    table.Range.ParagraphFormat.LineSpacingRule = c.wdLineSpaceSingle
    table.Range.ParagraphFormat.SpaceBefore = 0
    table.Range.ParagraphFormat.SpaceAfter = 0
    
    for i, row in enumerate(data):
        for j, cell_val in enumerate(row):
            cell = table.Cell(i + 1, j + 1)
            cell.Range.Text = cell_val
            if (i, j) in bold_cells:
                cell.Range.Font.Bold = True
            if (i, j) == (0, 0):
                placeholder = "___"
                cell.Range.Text = placeholder
                bm_start = cell.Range.Start
                bm_range = doc.Range(bm_start, bm_start + len(placeholder))
                doc.Bookmarks.Add("GuideName_2", bm_range)
            if (i, j) == (1, 0):
                placeholder = "___"
                cell.Range.Text = placeholder
                bm_start = cell.Range.Start
                bm_range = doc.Range(bm_start, bm_start + len(placeholder))
                doc.Bookmarks.Add("Designation_2", bm_range)
            if (i, j) == (0, 1):
                placeholder = "___"
                cell.Range.Text = placeholder
                bm_start = cell.Range.Start
                bm_range = doc.Range(bm_start, bm_start + len(placeholder))
                doc.Bookmarks.Add("Department_5", bm_range)
            if (i, j) == (2, 0):
                placeholder = "___"
                cell.Range.Text = placeholder + ","
                bm_start = cell.Range.Start
                bm_range = doc.Range(bm_start, bm_start + len(placeholder))
                doc.Bookmarks.Add("Department_6", bm_range)
            if (i, j) == (2, 1):
                placeholder = "___"
                cell.Range.Text = placeholder + ","
                bm_start = cell.Range.Start
                bm_range = doc.Range(bm_start, bm_start + len(placeholder))
                doc.Bookmarks.Add("Department_7", bm_range)

    for border_id in [
        c.wdBorderTop, c.wdBorderBottom, c.wdBorderLeft, c.wdBorderRight,
        c.wdBorderHorizontal, c.wdBorderVertical
    ]:
        border = table.Borders(border_id)
        border.LineStyle = c.wdLineStyleSingle
        border.Color = c.wdColorWhite

    cursor = table.Range.Duplicate
    cursor.Collapse(c.wdCollapseEnd)
    cursor.InsertParagraphAfter()
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()

    # time.sleep(0.1)
# _________________________________________________________________________________

    
    data = [
        ["", "Name", "Signature with Date"]
    ]

    bold_cells = [(0, 1), (0, 2)]

    rows = len(data)
    cols = max(len(row) for row in data)

    cursor.Collapse(c.wdCollapseEnd)
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1) 
    cursor.Select()

    table = doc.Tables.Add(cursor, NumRows=rows, NumColumns=cols)
    table.Range.Style = "Table Grid"
    
    table.Range.Font.Name = "Times New Roman"
    table.Range.Font.Size = 12
    table.Range.ParagraphFormat.Alignment = c.wdAlignParagraphCenter
    table.Range.ParagraphFormat.LineSpacingRule = c.wdLineSpaceSingle
    table.Range.ParagraphFormat.SpaceBefore = 0
    table.Range.ParagraphFormat.SpaceAfter = 0
    
    for i, row in enumerate(data):
        for j, cell_val in enumerate(row):
            cell = table.Cell(i + 1, j + 1)
            cell.Range.Text = cell_val
            if (i, j) in bold_cells:
                cell.Range.Font.Bold = True

    for border_id in [
        c.wdBorderTop, c.wdBorderBottom, c.wdBorderLeft, c.wdBorderRight,
        c.wdBorderHorizontal, c.wdBorderVertical
    ]:
        border = table.Borders(border_id)
        border.LineStyle = c.wdLineStyleSingle
        border.Color = c.wdColorWhite

    cursor = table.Range.Duplicate
    cursor.Collapse(c.wdCollapseEnd)
    cursor.InsertParagraphAfter()
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()

    # time.sleep(0.1)

# _________________________________________________________________________________
    
    data = [
        ["Examiner 1:", "", ""],
        ["Examiner 2:", "", ""]
    ]

    bold_cells = [(0, 0), (1, 0)]

    rows = len(data)
    cols = max(len(row) for row in data)

    cursor.Collapse(c.wdCollapseEnd)
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1) 
    cursor.Select()

    table = doc.Tables.Add(cursor, NumRows=rows, NumColumns=cols)
    table.Range.Style = "Table Grid"
    
    table.Range.Font.Name = "Times New Roman"
    table.Range.Font.Size = 12
    table.Range.ParagraphFormat.Alignment = c.wdAlignParagraphLeft
    table.Range.ParagraphFormat.LineSpacingRule = c.wdLineSpaceSingle
    table.Range.ParagraphFormat.SpaceBefore = 0
    table.Range.ParagraphFormat.SpaceAfter = 0
    
    for i, row in enumerate(data):
        for j, cell_val in enumerate(row):
            cell = table.Cell(i + 1, j + 1)
            cell.Range.Text = cell_val
            if (i, j) in bold_cells:
                cell.Range.Font.Bold = True

    for border_id in [
        c.wdBorderTop, c.wdBorderBottom, c.wdBorderLeft, c.wdBorderRight,
        c.wdBorderHorizontal, c.wdBorderVertical
    ]:
        border = table.Borders(border_id)
        border.LineStyle = c.wdLineStyleSingle
        border.Color = c.wdColorWhite

    cursor = table.Range.Duplicate
    cursor.Collapse(c.wdCollapseEnd)
    cursor.InsertParagraphAfter()
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()

    #insert_table(data, align = c.wdAlignParagraphLeft, bold_cells = bold_cells, transparent = True)
    # time.sleep(0.1)

# _________________________________________________________________________________
# _________________________________________________________________________________

    cursor.Collapse(c.wdCollapseEnd)
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1) 
    
    cursor.InsertBreak(c.wdPageBreak) 
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
    # time.sleep(0.1)
# _________________________________________________________________________________
# _________________________________________________________________________________

    word.Selection.ParagraphFormat.LineSpacingRule = c.wdLineSpace1pt5
    set_format(size=14, bold=True, align=c.wdAlignParagraphCenter, underline=c.wdUnderlineNone)
    word.Selection.TypeText("ACKNOWLEDGEMENT")
    word.Selection.TypeParagraph()

    set_format(size=12, bold=False, align=c.wdAlignParagraphJustify)
    word.Selection.TypeText("I take this opportunity to express my heartfelt gratitude to all those who supported and guided me throughout the development of this project, ")
    set_format(bold=True)
    add_bookmark("ProjectTitle_Ack", "___") 
    set_format(bold=False)
    word.Selection.TypeText(". Their contributions and encouragement were invaluable to the successful completion of this endeavour.")
    word.Selection.TypeParagraph()
    word.Selection.TypeParagraph()

    word.Selection.TypeText("First and foremost, I would like to extend my sincere thanks to the Dean of our institution, Prof. Eishwar N Maanay, for providing the resources and a conducive environment to undertake this project. Their constant support and emphasis on innovation inspired me to push my boundaries.")
    word.Selection.TypeParagraph()
    word.Selection.TypeParagraph()

    word.Selection.TypeText("I am immensely grateful to our Head of the Department, ")
    set_format(bold=True)
    add_bookmark("HODName_Ack", "___")
    set_format(bold=False)
    word.Selection.TypeText(", ")
    add_bookmark("Department_9", "___")
    word.Selection.TypeText(" for their unwavering support and guidance. Their insights and suggestions played a crucial role in shaping the direction of this project. Their encouragement throughout the process has been a source of great motivation.")
    word.Selection.TypeParagraph()
    word.Selection.TypeParagraph()

    word.Selection.TypeText("A special note of appreciation goes to my Guide, ")
    set_format(bold=True)
    add_bookmark("GuideName_Ack", "___")
    set_format(bold=False)
    word.Selection.TypeText(", ")
    add_bookmark("Designation_Ack", "___")
    word.Selection.TypeText(" for their technical expertise, and constructive feedback. Their patient guidance, timely advice, and constant encouragement helped me overcome challenges and refine the project to its current form.")
    word.Selection.TypeParagraph()
    word.Selection.TypeParagraph()

    word.Selection.TypeText("I also wish to express my deepest gratitude to my parents for their unconditional love, support, and encouragement throughout this journey. Their belief in my abilities has been my greatest strength, and their words of motivation have always driven me to excel.")
    word.Selection.TypeParagraph()
    word.Selection.TypeParagraph()

    word.Selection.TypeText("Lastly, I would like to thank my peers, friends, and everyone who contributed directly or indirectly to the successful completion of this project. Their encouragement and suggestions have been instrumental in making this project a success.")
    word.Selection.TypeParagraph()
    word.Selection.TypeParagraph()

    word.Selection.TypeText("This project would not have been possible without the collective support of everyone mentioned above. I am truly grateful for their contributions and look forward to utilizing the knowledge and skills gained from this experience in future endeavours.")
    # word.Selection.TypeParagraph() # Removed to prevent empty page
    
    # word.Selection.InsertParagraphAfter() # Avoid this if not needed
    word.Selection.InsertBreak(c.wdPageBreak)
    word.Selection.MoveLeft(Unit=1, Count=1)
    word.Selection.Delete(Unit=1, Count=1)
    word.Selection.MoveRight(Unit=1, Count=1)
    # cursor.Collapse(c.wdCollapseEnd)
    # cursor.Select()
    # time.sleep(0.1)
# _________________________________________________________________________________
# _________________________________________________________________________________

    set_format(size=14, bold=True, align=c.wdAlignParagraphCenter, underline=c.wdUnderlineNone)
    word.Selection.TypeText("ABSTRACT")
    word.Selection.TypeParagraph()

    word.Selection.ParagraphFormat.LineSpacingRule = c.wdLineSpace1pt5    
    set_format(size=12, bold=False, align=c.wdAlignParagraphJustify)
    add_bookmark("Abstract", "___")
    # time.sleep(0.1)
# _________________________________________________________________________________
# _________________________________________________________________________________

    cursor.Collapse(c.wdCollapseEnd)
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1) 
    
    cursor.InsertBreak(c.wdSectionBreakNextPage) 
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
    # time.sleep(0.1)
# _________________________________________________________________________________
# _________________________________________________________________________________

    sec = doc.Sections(2)  
    cursor = sec.Range.Duplicate
    cursor.Collapse(c.wdCollapseStart)
    cursor.Select()
    word.Selection.TypeParagraph()
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter
    cursor.Select()
    
    word.Selection.Font.Name = "Times New Roman"
    word.Selection.Font.Size = 14
    word.Selection.Font.Bold = True
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter
    word.Selection.TypeText("Table of Contents")
    word.Selection.TypeParagraph()

    data = [
        ["S.No", "Title", "Page No"],
        ["1", "___", "___"],
        ["2", "___", "___"],
        ["3", "___", "___"],
        ["4", "___", "___"],
        ["5", "___", "___"],
        ["6", "References", "___"],
    ]

    bold_cells = [(0, 0), (0, 1), (0, 2)]

    rows = len(data)
    cols = max(len(row) for row in data)

    cursor.Collapse(c.wdCollapseEnd)
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1) 
    cursor.Select()

    table = doc.Tables.Add(cursor, NumRows=rows, NumColumns=cols)
    table.Range.Style = "Table Grid"
    
    table.Range.Font.Name = "Times New Roman"
    table.Range.Font.Size = 12
    table.Range.ParagraphFormat.Alignment = c.wdAlignParagraphCenter
    table.Range.ParagraphFormat.LineSpacingRule = c.wdLineSpaceSingle
    table.Range.ParagraphFormat.SpaceBefore = 4
    table.Range.ParagraphFormat.SpaceAfter = 4
    
    table.Columns(1).SetWidth(cm_to_pt(1.25), c.wdAdjustNone)   
    table.Columns(2).SetWidth(cm_to_pt(13.75), c.wdAdjustNone)  
    table.Columns(3).SetWidth(cm_to_pt(2), c.wdAdjustNone) 
    
    for i, row in enumerate(data):
        for j, cell_val in enumerate(row):
            cell = table.Cell(i + 1, j + 1)
            cell.Range.Text = cell_val
            if (i, j) in bold_cells:
                cell.Range.Font.Bold = True
            if j == 1 and i > 0 and i < 6:
                placeholder = "___"
                cell.Range.Text = placeholder
                bm_start = cell.Range.Start
                bm_range = doc.Range(bm_start, bm_start + len(placeholder))
                doc.Bookmarks.Add(f"Chapter{i}Title", bm_range)
            if j == 2 and i > 0 and i < 6:
                placeholder = "___"
                cell.Range.Text = placeholder
                bm_start = cell.Range.Start
                bm_range = doc.Range(bm_start, bm_start + len(placeholder))
                doc.Bookmarks.Add(f"Chapter{i}Page", bm_range)
            if j == 2 and i == 6:
                placeholder = "___"
                cell.Range.Text = placeholder
                bm_start = cell.Range.Start
                bm_range = doc.Range(bm_start, bm_start + len(placeholder))
                doc.Bookmarks.Add("RefPage", bm_range)

    for border_id in [
        c.wdBorderTop, c.wdBorderBottom, c.wdBorderLeft, c.wdBorderRight,
        c.wdBorderHorizontal, c.wdBorderVertical
    ]:
        border = table.Borders(border_id)
        border.LineStyle = c.wdLineStyleSingle
        border.Color = c.wdColorBlack

    cursor = table.Range.Duplicate
    cursor.Collapse(c.wdCollapseEnd)
    cursor.InsertParagraphAfter()
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
    # time.sleep(0.1)

# _________________________________________________________________________________

    for i in range(1, 6):
        # ------------------------------------------
        # Section 1: Centered vertically (Title_2)
        # ------------------------------------------        
        cursor.Collapse(c.wdCollapseEnd)
        cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        cursor.InsertBreak(c.wdSectionBreakNextPage)
        cursor.Collapse(c.wdCollapseEnd)
        cursor.Select()

        word.Selection.Font.Name = "Times New Roman"
        set_format(size=16, bold=True, align=c.wdAlignParagraphCenter)

        center_pad_lines = 9
        for _ in range(center_pad_lines):
            word.Selection.TypeParagraph()
    
        # Title_2
        word.Selection.TypeText(f"Chapter {i}")
        word.Selection.TypeParagraph()
        placeholder = "___"
        word.Selection.TypeText(placeholder)
        bm_range = word.Selection.Range.Duplicate
        bm_start = bm_range.Start - len(placeholder)
        bm_range = doc.Range(bm_start, bm_start + len(placeholder))
        doc.Bookmarks.Add(f"Chapter{i}Title_2", bm_range)
        word.Selection.TypeParagraph()

        # -----------------------------------------------------
        # Section 2: Normal top alignment (Title_3 + Content)
        # -----------------------------------------------------
        cursor.Collapse(c.wdCollapseEnd)
        cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        cursor.InsertBreak(c.wdPageBreak)
        cursor.Collapse(c.wdCollapseEnd)
        cursor.Select()


        # Title_3
        placeholder = "___"
        word.Selection.TypeText(placeholder)
        bm_range = word.Selection.Range.Duplicate
        bm_start = bm_range.Start - len(placeholder)
        bm_range = doc.Range(bm_start, bm_start + len(placeholder))
        doc.Bookmarks.Add(f"Chapter{i}Title_3", bm_range)
        word.Selection.TypeParagraph()

        # Content
        word.Selection.ParagraphFormat.LineSpacingRule = c.wdLineSpace1pt5    
        word.Selection.Font.Size = 12
        word.Selection.Font.Bold = False
        word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphJustify

        placeholder = "___"
        word.Selection.TypeText(placeholder)
        content_range = word.Selection.Range.Duplicate  
        bm_start = content_range.Start - len(placeholder)
        content_bm_range = doc.Range(bm_start, bm_start + len(placeholder))
        doc.Bookmarks.Add(f"Chapter{i}Content", content_bm_range)
        word.Selection.TypeParagraph()


    # ---------------------------------------------
    # Final section break to isolate the next part
    # ---------------------------------------------
    cursor.Collapse(c.wdCollapseEnd)
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    cursor.InsertBreak(c.wdSectionBreakNextPage)
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
    # time.sleep(0.1)

# _________________________________________________________________________________
# _________________________________________________________________________________

    cursor.Collapse(c.wdCollapseEnd)
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    cursor.Select()
    
    word.Selection.Font.Name = "Times New Roman"                           
    word.Selection.Font.Size = 16                                          
    word.Selection.Font.Bold = True                                                
    word.Selection.Font.Italic = False                                       
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter     
    word.Selection.ParagraphFormat.LineSpacingRule = c.wdLineSpace1pt5    
    word.Selection.Font.Underline = c.wdUnderlineNone

    word.Selection.TypeText("REFERENCES")
    word.Selection.TypeParagraph()
    # time.sleep(0.1)
# _________________________________________________________________________________

    word.Selection.Font.Size = 12                                          
    word.Selection.Font.Bold = False                                                
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphJustify     
    word.Selection.Font.Underline = c.wdUnderlineNone

    placeholder = "___"
    word.Selection.TypeText(placeholder)
    bm_range = word.Selection.Range.Duplicate
    bm_start = bm_range.Start - len(placeholder)
    bm_range = doc.Range(bm_start, bm_start + len(placeholder))
    doc.Bookmarks.Add("References", bm_range)
    # time.sleep(0.1)
# _________________________________________________________________________________



# =================================================================================
    make_borders() # Call the function to set borders
    page_numbers() # Call the function to set page numbers

# _________________________________________________________________________________
# _________________________________________________________________________________


def make_borders():
    sec1 = doc.Sections(1) # Get the first section
    borders = sec1.Borders
    borders.DistanceFromTop = borders.DistanceFromBottom = 24
    borders.DistanceFromLeft = borders.DistanceFromRight = 12

    sec1.Range.Select() 
    word.Selection.Range.GoTo()

    for side in (c.wdBorderTop, c.wdBorderLeft, c.wdBorderBottom, c.wdBorderRight): # Set borders
        
        br = borders(side)
        br.LineStyle = c.wdLineStyleThinThickThinMedGap # Thin-Thick-Thin Medium Gap
        br.LineWidth = c.wdLineWidth300pt # 3 pt width
        br.Color = c.wdColorAutomatic # Automatic color (Black)

    # time.sleep(0.1)
# _________________________________________________________________________________
# _________________________________________________________________________________
def page_numbers():

    for idx, sec in enumerate(doc.Sections, start=1):
        sec.Range.InsertAfter("\r")
        if idx > 1:
            for hf_type in [c.wdHeaderFooterPrimary, c.wdHeaderFooterFirstPage]:
                sec.Footers(hf_type).LinkToPrevious = False
                sec.Headers(hf_type).LinkToPrevious = False

        if idx == 1 or idx == 2:
            for hf_type in [c.wdHeaderFooterPrimary, c.wdHeaderFooterFirstPage]:
                sec.Footers(hf_type).Range.Text = ""
                sec.Headers(hf_type).Range.Text = ""
            continue

        if idx == 3:
            sec.PageSetup.DifferentFirstPageHeaderFooter = False
            footer = sec.Footers(c.wdHeaderFooterPrimary)
            pnums = footer.PageNumbers
            pnums.RestartNumberingAtSection = True
            pnums.StartingNumber = 1
            pnums.Add(c.wdAlignParagraphCenter, False)

        if idx >= 4 and idx < 8:
            sec.PageSetup.DifferentFirstPageHeaderFooter = True
            pfooter = sec.Footers(c.wdHeaderFooterPrimary)
            ppnums = pfooter.PageNumbers
            ppnums.RestartNumberingAtSection = False
            ppnums.Add(c.wdAlignParagraphCenter, False)

            sec.Footers(c.wdHeaderFooterFirstPage).Range.Text = ""


# _________________________________________________________________________________
# _________________________________________________________________________________


# ---------------------------------------------------------------------------------
# def replace_bookmarks(data_dict: dict):
#     """
#     Replaces bookmarks in the Word document with values from a dictionary.
#     This function iterates through the provided dictionary and checks if each key exists as a bookmark in the document.
#     If a bookmark exists, it replaces the text of that bookmark with the corresponding value from the dictionary.

#     :param data_dict: A dictionary where keys are bookmark names and values are the text to replace them with.
#     :type data_dict: dict
#     """
    
#     all_bm_names = [bm.Name for bm in doc.Bookmarks] # Get all bookmark names in the document

#     # Only these exact bookmarks will get a newline after replacement
#     newline_bookmark_names = {
#         "ProjectTitle", "NameAndUSN", "GuideName",
#         "Chapter1Title", "Chapter2Title", "Chapter3Title", "Chapter4Title", "Chapter5Title",
#         "Chapter1Content", "Chapter2Content", "Chapter3Content", "Chapter4Content", "Chapter5Content"
#     }
    
#     rebookmarks = [] # To store bookmarks that are re-added after replacement
    
#     for key, value in data_dict.items(): # Iterate through the dictionary
#         matching_bms = [bm for bm in all_bm_names if bm.startswith(key)] # Find all bookmarks that start with the key
#         if not matching_bms:
#             continue # If no matching bookmarks found, skip to the next key
        
#         for name in matching_bms:
#             if not doc.Bookmarks.Exists(name):
#                 continue # If the bookmark does not exist, skip
            
#             bm_range = doc.Bookmarks(name).Range # range of bookmark
#             bm_start = bm_range.Start # start position of bookmark
#             add_newline = name in newline_bookmark_names # Check if this bookmark should have a newline after it
#             insert_text = value + ("\n" if add_newline else " ") # Replace bookmark text with value
            
#             bm_range.Text = insert_text # Replace bookmark text with value
            
#             new_range = doc.Range(bm_start, bm_start + len(insert_text)) # create a new range for the bookmark
#             rebookmarks.append((name, new_range)) # Store the bookmark name and new range
            
#             new_range.Select() # Select the new range
#             word.ActiveWindow.ScrollIntoView(word.Selection.Range, True) # Scroll to the new range

#         for name, rng in rebookmarks: # Re-add the bookmarks with the new ranges
#             try:
#                 doc.Bookmarks.Add(name, rng)
#             except:
#                 print(f"⚠️ Could not re-add bookmark: {name}")

#     title = data_dict.get("ProjectTitle")
#     year = data_dict.get("Year")

#     if title or year:
#         for idx, section in enumerate(doc.Sections, start=1):
#             if idx == 1 or idx == 2:
#                 continue

#             # HEADER: Left-align project title
#             header = section.Headers(c.wdHeaderFooterPrimary)
#             header.LinkToPrevious = False
#             if title:
#                 header.Range.Text = title
#                 header.Range.ParagraphFormat.Alignment = c.wdAlignParagraphLeft

#             # FOOTER: Left = dept, Center = year, Right = page number
#             footer = section.Footers(c.wdHeaderFooterPrimary)
#             footer.LinkToPrevious = False
#             rng = footer.Range
#             rng.Text = ""

#             table = rng.Tables.Add(rng, NumRows=1, NumColumns=3)
#             table.PreferredWidthType = c.wdPreferredWidthPercent
#             table.PreferredWidth = 100
#             table.Borders.Enable = False

#             # Left = Dept.
#             table.Cell(1, 1).Range.Text = "Dept. of CSE, BNMIT"
#             table.Cell(1, 1).Range.ParagraphFormat.Alignment = c.wdAlignParagraphLeft

#             # Center = Year (only if provided)
#             if year:
#                 table.Cell(1, 2).Range.Text = year
#             table.Cell(1, 2).Range.ParagraphFormat.Alignment = c.wdAlignParagraphCenter

#             # Right = Page number
#             right_range = table.Cell(1, 3).Range
#             right_range.Collapse(c.wdCollapseStart)
#             right_range.Fields.Add(right_range, c.wdFieldPage)
#             right_range.ParagraphFormat.Alignment = c.wdAlignParagraphRight


def replace_bookmarks(data_dict: dict):
    """
    Replaces bookmarks in the Word document with values from a dictionary.
    Also inserts images after Chapter{i}Content bookmarks if matching files are found.
    """
    transformed_data = {}
    
    dept_short_forms = {
        "COMPUTER SCIENCE AND ENGINEERING": "Dept. of CSE",
        "ELECTRICAL AND COMMUNICATION ENGINEERING": "Dept. of ECE",
        "INFORMATION SCIENCE AND ENGINEERING": "Dept. of ISE",
        "MECHANICAL ENGINEERING": "Dept. of ME",
        "CIVIL ENGINEERING": "Dept. of CE",
        "ELECTRONICS AND INSTRUMENTATION ENGINEERING": "Dept. of EIE",
        "ARTIFICIAL INTELLIGENCE AND MACHINE LEARNING": "Dept. of AIML",
        "ELECTRICAL AND ELECTRONICS ENGINEERING": "Dept. of EEE"
    }
    
    hod_titles = {
        "COMPUTER SCIENCE AND ENGINEERING": "Dr. Chayadevi M.L",
        "ELECTRICAL AND COMMUNICATION ENGINEERING": "Dr. P. A. Vijaya", 
        "INFORMATION SCIENCE AND ENGINEERING": "Dr. S. Srividhya",
        "MECHANICAL ENGINEERING": "Dr. B.S. Anil Kumar",
        "CIVIL ENGINEERING": "Dr. S.B. Anadinni",
        "ELECTRONICS AND INSTRUMENTATION ENGINEERING": "Dr. K.S. Jyothi",
        "ARTIFICIAL INTELLIGENCE AND MACHINE LEARNING": "Dr. Saritha Chakrasali",
        "ELECTRICAL AND ELECTRONICS ENGINEERING": "Dr. R.V. Parimala"
    }

    department_value = data_dict.get("Department", "").strip()

    # Apply transformed values based on that single input
    if department_value:
        # HOD full name → for Department_5
        hod_value = hod_titles.get(department_value, department_value)
        transformed_data["Department_5"] = hod_value
        # Department_8 is used in "Department of [Department_8]". Should be full name or just branch.
        # User requested: "department of computer science and engineering not department hod name"
        transformed_data["Department_8"] = department_value 

        # Short form dept → for Department_6 and Department_7
        short_form = dept_short_forms.get(department_value, department_value)
        transformed_data["Department_6"] = short_form
        transformed_data["Department_7"] = short_form
        transformed_data["Department_9"] = department_value # Changed to Full Name for Acknowledgement
        
        # User: "We also express our sincere thanks to all the staff members of the Department of ___"
        # Should be full name or short form? "Department of Dept. of CSE" is weird.
        # "Department of Computer Science and Engineering" is better.
        transformed_data["Department_10"] = department_value
        
        # Explicit mappings for Title Page and Certificate where raw 'Department' was missing
        transformed_data["Department"] = department_value    # Title Page: "In [Department]"
        transformed_data["Department_4"] = department_value  # Certificate: "Bachelor of Engineering in [Department_4]"
        
        # For Acknowledgement HOD Name
        transformed_data["HODName_Ack"] = hod_value
        
        # New Acknowledgement Mappings
        transformed_data["ProjectTitle_Ack"] = data_dict.get("ProjectTitle", "")
        transformed_data["GuideName_Ack"] = data_dict.get("GuideName", "")
        transformed_data["Designation_Ack"] = data_dict.get("Designation", "")
        
        # New Acknowledgement Mappings
        transformed_data["ProjectTitle_Ack"] = data_dict.get("ProjectTitle", "")
        transformed_data["GuideName_Ack"] = data_dict.get("GuideName", "")
        transformed_data["Designation_Ack"] = data_dict.get("Designation", "")

    # Also carry over other keys from data_dict directly
    for key, value in data_dict.items():
        if key != "Department":  # Already handled separately
            if key == "NameAndUSN":
                # Special handling for Certificate Page usage
                # If NameAndUSN has newlines (from multiline input), replace them with commas for the inline certificate version
                inline_names = value.replace("\n", ", ")
                transformed_data["NameAndUSN_2"] = inline_names
            
            transformed_data[key] = value
            
    all_bm_names = [bm.Name for bm in doc.Bookmarks]  # Get all bookmark names in the document

    # These bookmarks should have a newline after the inserted value
    newline_bookmark_names = {
        "ProjectTitle", "NameAndUSN", "GuideName", "Designation",
        "Department_2", "Department_3",
        "Chapter1Title", "Chapter2Title", "Chapter3Title", "Chapter4Title", "Chapter5Title",
        "Chapter1Content", "Chapter2Content", "Chapter3Content", "Chapter4Content", "Chapter5Content"
    }

    rebookmarks = []  # To store bookmarks that need to be re-added after replacement

    # MAIN REPLACEMENT LOOP - Uses transformed_data to ensure derived keys are covered
    for key, value in transformed_data.items():
        matching_bms = [bm for bm in all_bm_names if bm.startswith(key)]
        if not matching_bms:
            continue

        for name in matching_bms:
            # Skip if this specific bookmark name doesn't exist 
            if not doc.Bookmarks.Exists(name):
                continue

            # CRITICAL: Prevent "NameAndUSN" key from overwriting "NameAndUSN_2" bookmark
            # if "NameAndUSN_2" has its own entry in transformed_data.
            if name != key and name in transformed_data:
                continue 
            
            bm_range = doc.Bookmarks(name).Range
            bm_start = bm_range.Start
            
            add_newline = name in newline_bookmark_names
            insert_text = value + ("\n" if add_newline else "") # Removed space for inline bookmarks
            
            bm_range.Text = insert_text
            
            new_range = doc.Range(bm_start, bm_start + len(insert_text))
            rebookmarks.append((name, new_range))
            
            new_range.Select()
            word.ActiveWindow.ScrollIntoView(word.Selection.Range, True)
            
            # --- Handle images (ChapterContent logic) ---
            chapter_match = re.match(r"Chapter(\d)Content", name)
            if chapter_match:
                chapter_num = int(chapter_match.group(1))

                def extract_figure_index(p):
                    match = re.search(rf"Fig {chapter_num}\.(\d+)", p.stem)
                    if match:
                        return float(match.group(1))
                    return float('inf')

                image_files = sorted(
                    ASSET_DIR.glob(f"Fig {chapter_num}.*"),
                    key=extract_figure_index
                )

                if image_files:
                    # Step 1: Define start of insertion range
                    chapter_end = new_range.End

                    # Step 2: Define end of chapter by checking next chapter title
                    next_title = f"Chapter{chapter_num + 1}Title_2"
                    if next_title in [b.Name for b in doc.Bookmarks]:
                        chapter_limit = doc.Bookmarks(next_title).Range.Start
                    else:
                        chapter_limit = doc.Content.End

                    # Step 3: Define range to check for existing figure captions
                    safe_start = min(chapter_end, chapter_limit)
                    safe_end = max(chapter_end, chapter_limit)
                    if safe_end > doc.Content.End:
                        safe_end = doc.Content.End

                    scan_range = doc.Range(safe_start, safe_end)
                    existing_text = scan_range.Text

                    # Step 4: Begin inserting images in order using a safe advancing range
                    insert_range = doc.Range(chapter_end, chapter_end)
                    insert_range.Collapse(c.wdCollapseStart)

                    for img in image_files:
                        fig_index = img.stem.split('.')[-1]
                        fig_label = f"Fig {chapter_num}.{fig_index}"

                        if fig_label in existing_text:
                            continue  # Already inserted

                        # Step 1: Remember where image is being inserted
                        image_start = insert_range.Start
                        
                        # --- Smart Placement Logic ---
                        # 1. Calc target dimensions
                        # Word restricts images to page margins. Assume max width 450pt (approx 16cm).
                        max_width_pt = 450 
                        with Image.open(str(img.resolve())) as pil_img:
                             w_px, h_px = pil_img.size
                             aspect = h_px / w_px
                             
                             # Convert px to pt (Approximate: 1 px = 0.75 pt at 96 DPI)
                             # This estimates the "Natural" size Word will use.
                             natural_width_pt = w_px * 0.75
                             
                             # If natural width > max page width, it shrinks. Else it stays natural.
                             effective_width_pt = min(natural_width_pt, max_width_pt)
                             
                             target_height_pt = effective_width_pt * aspect 
                        
                        # 2. Check available space
                        # Get current vertical position
                        try:
                            wdVerticalPositionRelativeToPage = 6 # Constant
                            current_vertical_pos = insert_range.Information(wdVerticalPositionRelativeToPage)
                            
                            # Get Page Height and Margin
                            page_height = doc.PageSetup.PageHeight
                            bottom_margin = doc.PageSetup.BottomMargin
                            limit = page_height - bottom_margin
                            
                            available_space = limit - current_vertical_pos
                            caption_buffer = 60 # Points for caption + spacing
                            
                            # 3. Decide on Page Break
                            # If the image WOULD fit if shrunk, but we aren't forcing shrink, 
                            # checking against 'max possible height' is safer to prevent overflow.
                            if (current_vertical_pos + target_height_pt + caption_buffer) > limit:
                                # Not enough space, force page break
                                insert_range.InsertBreak(c.wdPageBreak)
                                # Update range after break
                                insert_range.Collapse(c.wdCollapseEnd)
                                
                        except Exception as e:
                            print(f"⚠️ Calculation error: {e}. Letting Word decide placement.")
                        
                        # Step 2: Insert image
                        # Use a dedicated range for image insertion to avoid style bleed
                        img_range = insert_range.Duplicate
                        img_shape = img_range.InlineShapes.AddPicture(str(img.resolve()), LinkToFile=False, SaveWithDocument=True)
                        
                        # Remove explicit resizing to respect user request
                        # img_shape.Width = target_width_pt 
                        
                        # Center the image
                        img_shape.Range.ParagraphFormat.Alignment = c.wdAlignParagraphCenter
                        img_shape.Range.ParagraphFormat.KeepWithNext = True # Keep image with its caption
                        
                        # Step 3: Insert Caption
                        # Move to end of image shape itself to guarantee order
                        caption_range = img_shape.Range.Duplicate
                        caption_range.Collapse(c.wdCollapseEnd)
                        caption_range.InsertParagraphAfter()
                        caption_range.Collapse(c.wdCollapseEnd)
                        
                        caption_range.Text = fig_label
                        # Explicitly reset formatting for caption to avoid inheriting Title styles
                        caption_range.Font.Name = "Times New Roman"
                        caption_range.Font.Size = 12
                        caption_range.Font.Bold = False
                        caption_range.ParagraphFormat.Alignment = c.wdAlignParagraphCenter
                        caption_range.ParagraphFormat.SpaceAfter = 12 # Give some breathing room
                        
                        caption_range.InsertParagraphAfter()
                        
                        # Step 4: Advance safely
                        insert_range = caption_range.Duplicate
                        insert_range.Collapse(c.wdCollapseEnd)

    # --- Re-add bookmarks ---
    for name, rng in rebookmarks:
        try:
            doc.Bookmarks.Add(name, rng)
        except:
            print(f"⚠️ Could not re-add bookmark: {name}")

    # --- Header/Footer logic ---
    title = data_dict.get("ProjectTitle")
    year = data_dict.get("Year")

    if title or year:
        for idx, section in enumerate(doc.Sections, start=1):
            if idx == 1 or idx == 2:
                continue

            # HEADER: Left-align project title
            if idx > 1:
                header = section.Headers(c.wdHeaderFooterPrimary)
                header.LinkToPrevious = False
                if title:
                    header.Range.Text = title
                    header.Range.ParagraphFormat.Alignment = c.wdAlignParagraphLeft

                # FOOTER: Left = dept, Center = year, Right = page number
                footer = section.Footers(c.wdHeaderFooterPrimary)
                footer.LinkToPrevious = False
                rng = footer.Range
                rng.Text = ""

                table = rng.Tables.Add(rng, NumRows=1, NumColumns=3)
                table.PreferredWidthType = c.wdPreferredWidthPercent
                table.PreferredWidth = 100
                table.Borders.Enable = False

                # Left = Dept.
                table.Cell(1, 1).Range.Text = "Dept. of CSE, BNMIT"
                table.Cell(1, 1).Range.ParagraphFormat.Alignment = c.wdAlignParagraphLeft

                # Center = Year (only if provided)
                if year:
                    table.Cell(1, 2).Range.Text = year
                table.Cell(1, 2).Range.ParagraphFormat.Alignment = c.wdAlignParagraphCenter

                # Right = Page number
                right_range = table.Cell(1, 3).Range
                right_range.Collapse(c.wdCollapseStart)
                right_range.Fields.Add(right_range, c.wdFieldPage)
                right_range.ParagraphFormat.Alignment = c.wdAlignParagraphRight


            
# ---------------------------------------------------------------------------------

def update_index_page_numbers():
    # Attempt to use wdActiveEndAdjustedPageNumber (4) for restart-aware numbering
    # If not in constants, define it manually
    wdActiveEndAdjustedPageNumber = getattr(c, 'wdActiveEndAdjustedPageNumber', 4)

    for i in range(1, 6):
        title_bm = f"Chapter{i}Title_2"
        page_bm = f"Chapter{i}Page"  # This is in the index table
        
        if doc.Bookmarks.Exists(title_bm) and doc.Bookmarks.Exists(page_bm):
            title_range = doc.Bookmarks(title_bm).Range
            # Use AdjustedPageNumber to respect the footer restart
            page_number = title_range.Information(wdActiveEndAdjustedPageNumber)

            # Replace the index placeholder bookmark with the actual page number
            bm_range = doc.Bookmarks(page_bm).Range
            bm_start = bm_range.Start
            bm_range.Text = str(page_number) # No static offset needed now

            # Re-bookmark the range so that the bookmark persists
            new_range = doc.Range(bm_start, bm_start + len(str(page_number)))
            try:
                doc.Bookmarks.Add(page_bm, new_range)
            except:
                print(f"⚠️ Could not re-add bookmark: {page_bm}")
                
        if doc.Bookmarks.Exists("References") and doc.Bookmarks.Exists("RefPage"):
            ref_range = doc.Bookmarks("References").Range
            ref_page = ref_range.Information(wdActiveEndAdjustedPageNumber) 

            bm_range = doc.Bookmarks("RefPage").Range
            bm_start = bm_range.Start
            bm_range.Text = str(ref_page)

            # Re-bookmark the range so that the bookmark persists
            new_range = doc.Range(bm_start, bm_start + len(str(ref_page)))
            try:
                doc.Bookmarks.Add("RefPage", new_range)
            except:
                print(f"⚠️ Could not re-add bookmark: RefPage")

# ================================================================================= 
# =================================================================================

def save_document():
    """
    Saves the current Word document to the specified path.
    """
    update_index_page_numbers()
    doc.Fields.Update()
    for field in doc.Fields:
        field.Update()
    for section in doc.Sections:
        section.Headers(c.wdHeaderFooterPrimary).Range.Fields.Update()
        section.Footers(c.wdHeaderFooterPrimary).Range.Fields.Update()
    doc.SaveAs(str(DOC_PATH), FileFormat=c.wdFormatDocumentDefault)
    CTkMessagebox(title="Saved", message=f"The report has been successfully saved.\n\nSave Location: {DOC_PATH.resolve()}", icon="check")
    
# ================================================================================= 
# =================================================================================