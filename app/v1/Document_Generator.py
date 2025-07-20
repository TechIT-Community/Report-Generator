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
    
    word.Selection.Font.Name = "Times New Roman"                            # Font Name
    word.Selection.Font.Size = 15                                           # Font Size
    word.Selection.Font.Bold = True                                         # Bold        
    word.Selection.Font.Italic = False                                      # Italic  
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter     # Alignment
    word.Selection.ParagraphFormat.LineSpacingRule = c.wdLineSpaceSingle    # Line Spacing
    word.Selection.Font.Underline = c.wdUnderlineNone                       # No Underline


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
    cursor.InsertParagraphAfter() # Insert a paragraph break after the image
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
    # time.sleep(0.1)
# _________________________________________________________________________________

    word.Selection.Font.Size = 11
    word.Selection.TypeText("A MINI PROJECT\vOn")
    word.Selection.TypeParagraph()
    # time.sleep(0.1)
# _________________________________________________________________________________
    
    word.Selection.Font.Size = 15
    placeholder = "___\n"
    word.Selection.TypeText(placeholder)
    bm_range = word.Selection.Range.Duplicate
    bm_start = bm_range.Start - len(placeholder)
    bm_range = doc.Range(bm_start, bm_start + len(placeholder))
    doc.Bookmarks.Add("ProjectTitle", bm_range)
    # time.sleep(0.1)
# _________________________________________________________________________________

    word.Selection.Font.Size = 11
    word.Selection.Font.Bold = False
    word.Selection.Font.Italic = True
    word.Selection.TypeText("Submitted in partial fulfilment of the requirements for the award of degree")
    word.Selection.TypeParagraph()
    # time.sleep(0.1)
# _________________________________________________________________________________

    word.Selection.Font.Name = "Times New Roman"
    word.Selection.Font.Bold = False
    word.Selection.Font.Italic = False
    word.Selection.TypeText("Bachelor of Engineering\vIn\v")
    # time.sleep(0.1)

    word.Selection.Font.Bold = True
    word.Selection.TypeText("Computer Science and Engineering")
    word.Selection.TypeParagraph()    

    word.Selection.Font.Bold = False
    word.Selection.TypeText("Submitted by")
    word.Selection.TypeParagraph()    
    # time.sleep(0.1)
# _________________________________________________________________________________

    word.Selection.Font.Bold = True
    placeholder = "___\n"
    word.Selection.TypeText(placeholder)
    bm_range = word.Selection.Range.Duplicate
    bm_start = bm_range.Start - len(placeholder)
    bm_range = doc.Range(bm_start, bm_start + len(placeholder))
    doc.Bookmarks.Add("NameAndUSN", bm_range)
    # time.sleep(0.1)
# _________________________________________________________________________________

    word.Selection.Font.Bold = False
    word.Selection.TypeText("Under the guidance of\v")    
    # time.sleep(0.1)
# _________________________________________________________________________________
    
    word.Selection.Font.Bold = True
    placeholder = "___\n"
    word.Selection.TypeText(placeholder)
    bm_range = word.Selection.Range.Duplicate
    bm_start = bm_range.Start - len(placeholder)
    bm_range = doc.Range(bm_start, bm_start + len(placeholder))
    doc.Bookmarks.Add("GuideName", bm_range)
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
    cursor.InsertParagraphAfter()
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
    # time.sleep(0.1)
# _________________________________________________________________________________

    word.Selection.Font.Bold = True
    word.Selection.TypeText("DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING")
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
    cursor.InsertParagraphAfter()
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

    word.Selection.TypeText("DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING")
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
    # time.sleep(0.1)

# _________________________________________________________________________________
    
    placeholder = "___ "
    word.Selection.TypeText(placeholder)
    bm_range = word.Selection.Range.Duplicate
    bm_start = bm_range.Start - len(placeholder)
    bm_range = doc.Range(bm_start, bm_start + len(placeholder))
    doc.Bookmarks.Add("ProjectTitle_2", bm_range)
    # time.sleep(0.1)
# _________________________________________________________________________________

    word.Selection.TypeText("has been successfully completed and is a bonafide work carried out by ")
    # time.sleep(0.1)
# _________________________________________________________________________________

    placeholder = "___ "
    word.Selection.TypeText(placeholder)
    bm_range = word.Selection.Range.Duplicate
    bm_start = bm_range.Start - len(placeholder)
    bm_range = doc.Range(bm_start, bm_start + len(placeholder))
    doc.Bookmarks.Add("NameUSN", bm_range)
    # time.sleep(0.1)
# _________________________________________________________________________________

    word.Selection.TypeText("bonafide students of ")
    # time.sleep(0.1)
# _________________________________________________________________________________

    placeholder = "___ "
    word.Selection.TypeText(placeholder)
    bm_range = word.Selection.Range.Duplicate
    bm_start = bm_range.Start - len(placeholder)
    bm_range = doc.Range(bm_start, bm_start + len(placeholder))
    doc.Bookmarks.Add("Sem", bm_range)
    # time.sleep(0.1)
# _________________________________________________________________________________

    word.Selection.TypeText("Semester B.E., B.N.M. Institute of Technology, an Autonomous Institution "
                            "under Visvesvaraya Technological University, Belagavi submitted in partial "
                            "fulfilment for the award of "
    )
    word.Selection.Font.Bold = True
    word.Selection.TypeText("Bachelor of Engineering in COMPUTER SCIENCE AND ENGINEERING, ")
    word.Selection.Font.Bold = False
    
    word.Selection.TypeText("during the year ")
    # time.sleep(0.1)
# _________________________________________________________________________________
    
    placeholder = "___ "
    word.Selection.TypeText(placeholder)
    bm_range = word.Selection.Range.Duplicate
    bm_start = bm_range.Start - len(placeholder)
    bm_range = doc.Range(bm_start, bm_start + len(placeholder))
    doc.Bookmarks.Add("Year", bm_range)
    # time.sleep(0.1)
# _________________________________________________________________________________

    word.Selection.TypeText("It is certified that all corrections / suggestions indicated for Internal Assessment "
                            "have been incorporated in the Report. The report has been approved as it satisfied "
                            "the academic requirements in respect project work prescribed by the said degree. ")
    # time.sleep(0.1)
# _________________________________________________________________________________

    data = [
        ["___",     "Dr. Chayadevi M L", "Dr. S Y Kulkarni"],
        ["Professor,",       "Professor and HOD,", "Additional Director"],
        ["Dept. of CSE,",     "Dept. of CSE,",      "and Principal,"],
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
                
                # Place bookmark over the exact range
                bm_start = cell.Range.Start
                bm_range = doc.Range(bm_start, bm_start + len(placeholder))
                doc.Bookmarks.Add("GuideName_2", bm_range)

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

    word.Selection.Font.Size = 16
    word.Selection.Font.Bold = True
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter
    word.Selection.TypeText("ACKNOWLEDGEMENT")
    word.Selection.TypeParagraph()
    # time.sleep(0.1)
# _________________________________________________________________________________

    word.Selection.Font.Size = 12
    word.Selection.Font.Bold = False
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphJustify
    word.Selection.TypeText(
        "I take this opportunity to express my heartfelt gratitude to all those who supported and guided me "
        "throughout the development of this project, the Automatic License Plate Recognition System. Their "
        "contributions and encouragement were invaluable to the successful completion of this endeavour.\n\n"
        
        "First and foremost, I would like to extend my sincere thanks to the Dean of our institution, Prof. "
        "Eishwar N Maanay, for providing the resources and a conducive environment to undertake this project. "
        "Their constant support and emphasis on innovation inspired me to push my boundaries.\n\n"

        "I am immensely grateful to our Head of the Department, Dr. Chayadevi M.L, Professor, Dept. of CSE, "
        "for their unwavering support and guidance. Their insights and suggestions played a crucial role in shaping "
        "the direction of this project. Their encouragement throughout the process has been a source of great motivation.\n\n"

        "A special note of appreciation goes to my Guide, Dr. Anitha N, Professor, for her invaluable mentorship, "
        "technical expertise, and constructive feedback. Their patient guidance, timely advice, and constant "
        "encouragement helped me overcome challenges and refine the project to its current form.\n\n"

        "I also wish to express my deepest gratitude to my parents for their unconditional love, support, and "
        "encouragement throughout this journey. Their belief in my abilities has been my greatest strength, "
        "and their words of motivation have always driven me to excel.\n\n"

        "Lastly, I would like to thank my peers, friends, and everyone who contributed directly or indirectly to the successful "
        "completion of this project. Your encouragement and suggestions have   been instrumental in making this project a success.\n\n"

        "This project would not have been possible without the collective support of everyone mentioned above. I am truly grateful "
        "for their contributions and look forward to utilizing the knowledge and skills gained from this experience in future endeavours.\n"
    )
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

    word.Selection.Font.Name = "Times New Roman"                           
    word.Selection.Font.Size = 16                                          
    word.Selection.Font.Bold = True                                                
    word.Selection.Font.Italic = False                                       
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter     
    word.Selection.ParagraphFormat.LineSpacingRule = c.wdLineSpace1pt5    
    word.Selection.Font.Underline = c.wdUnderlineNone

    word.Selection.TypeText("ABSTRACT")
    word.Selection.TypeParagraph()
    # time.sleep(0.1)
# _________________________________________________________________________________

    word.Selection.Font.Size = 12                                          
    word.Selection.Font.Bold = False                                                
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphJustify     
    word.Selection.ParagraphFormat.LineSpacingRule = c.wdLineSpace1pt5    

    placeholder = "___"
    word.Selection.TypeText(placeholder)
    bm_range = word.Selection.Range.Duplicate
    bm_start = bm_range.Start - len(placeholder)
    bm_range = doc.Range(bm_start, bm_start + len(placeholder))
    doc.Bookmarks.Add("Abstract", bm_range)
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
        word.Selection.Font.Size = 16
        word.Selection.Font.Bold = True
        word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter

        center_pad_lines = 10
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


        placeholder = "___"
        word.Selection.TypeText(placeholder)
        bm_range = word.Selection.Range.Duplicate
        bm_start = bm_range.Start - len(placeholder)
        bm_range = doc.Range(bm_start, bm_start + len(placeholder))
        doc.Bookmarks.Add(f"Chapter{i}Title_3", bm_range)
        word.Selection.TypeParagraph()

        # Content
        word.Selection.Font.Size = 12
        word.Selection.Font.Bold = False
        word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphJustify

        placeholder = "___"
        word.Selection.TypeText(placeholder)
        bm_range = word.Selection.Range.Duplicate
        bm_start = bm_range.Start - len(placeholder)
        bm_range = doc.Range(bm_start, bm_start + len(placeholder))
        doc.Bookmarks.Add(f"Chapter{i}Content", bm_range)
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
    word.Selection.Font.Underline = c.wdUnderlineSingle

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
        for hf_type in [c.wdHeaderFooterPrimary, c.wdHeaderFooterFirstPage]:
            sec.Footers(hf_type).LinkToPrevious = False

        if idx == 1 or idx == 2:
            for hf_type in [c.wdHeaderFooterPrimary, c.wdHeaderFooterFirstPage]:
                sec.Footers(hf_type).Range.Text = ""
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
def replace_bookmarks(data_dict: dict):
    """
    Replaces bookmarks in the Word document with values from a dictionary.
    This function iterates through the provided dictionary and checks if each key exists as a bookmark in the document.
    If a bookmark exists, it replaces the text of that bookmark with the corresponding value from the dictionary.

    :param data_dict: A dictionary where keys are bookmark names and values are the text to replace them with.
    :type data_dict: dict
    """
    
    all_bm_names = [bm.Name for bm in doc.Bookmarks] # Get all bookmark names in the document

    # Only these exact bookmarks will get a newline after replacement
    newline_bookmark_names = {
        "ProjectTitle", "NameAndUSN", "GuideName",
        "Chapter1Title", "Chapter2Title", "Chapter3Title", "Chapter4Title", "Chapter5Title",
        "Chapter1Content", "Chapter2Content", "Chapter3Content", "Chapter4Content", "Chapter5Content"
    }
    
    rebookmarks = [] # To store bookmarks that are re-added after replacement
    
    for key, value in data_dict.items(): # Iterate through the dictionary
        matching_bms = [bm for bm in all_bm_names if bm.startswith(key)] # Find all bookmarks that start with the key
        if not matching_bms:
            continue # If no matching bookmarks found, skip to the next key
        
        for name in matching_bms:
            if not doc.Bookmarks.Exists(name):
                continue # If the bookmark does not exist, skip
            
            bm_range = doc.Bookmarks(name).Range # range of bookmark
            bm_start = bm_range.Start # start position of bookmark
            add_newline = name in newline_bookmark_names # Check if this bookmark should have a newline after it
            insert_text = value + ("\n" if add_newline else " ") # Replace bookmark text with value
            
            bm_range.Text = insert_text # Replace bookmark text with value
            
            new_range = doc.Range(bm_start, bm_start + len(insert_text)) # create a new range for the bookmark
            rebookmarks.append((name, new_range)) # Store the bookmark name and new range
            
            new_range.Select() # Select the new range
            word.ActiveWindow.ScrollIntoView(word.Selection.Range, True) # Scroll to the new range

        for name, rng in rebookmarks: # Re-add the bookmarks with the new ranges
            try:
                doc.Bookmarks.Add(name, rng)
            except:
                print(f"⚠️ Could not re-add bookmark: {name}")
            
# ---------------------------------------------------------------------------------

def update_index_page_numbers():
    for i in range(1, 6):
        title_bm = f"Chapter{i}Title_2"
        page_bm = f"Chapter{i}Page"  # This is in the index table
        offset = 5
        
        if doc.Bookmarks.Exists(title_bm) and doc.Bookmarks.Exists(page_bm):
            title_range = doc.Bookmarks(title_bm).Range
            page_number = title_range.Information(c.wdActiveEndPageNumber)

            # Replace the index placeholder bookmark with the actual page number
            bm_range = doc.Bookmarks(page_bm).Range
            bm_start = bm_range.Start
            bm_range.Text = str(page_number - offset)

            # Re-bookmark the range so that the bookmark persists
            new_range = doc.Range(bm_start, bm_start + len(str(page_number)))
            try:
                doc.Bookmarks.Add(page_bm, new_range)
            except:
                print(f"⚠️ Could not re-add bookmark: {page_bm}")
                
        if doc.Bookmarks.Exists("References") and doc.Bookmarks.Exists("RefPage"):
            ref_range = doc.Bookmarks("References").Range
            ref_page = ref_range.Information(c.wdActiveEndPageNumber) 

            bm_range = doc.Bookmarks("RefPage").Range
            bm_start = bm_range.Start
            bm_range.Text = str(ref_page - offset)

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
    print("✅ Saved:", DOC_PATH.resolve())
    
# ================================================================================= 
# =================================================================================