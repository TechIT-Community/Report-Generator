"""
Imports for opening and manipulating Word using macros, findind paths to save and time for waiting.    
"""
import win32com.client as win32 # Interacting with word
from win32com.client import constants as c # All constants from win32com.client for word
from pathlib import Path # For file paths
import win32gui # For GUI like showing the window
import win32con # For window constanrs like SW_RESTORE
import time # For waiting for a while before executing next command

# =================================================================================

# Helping Functions
cm_to_pt = lambda cm: cm * 28.3464566929133858 # doc uses points, 1 cm = 28.35 points

# =================================================================================

# Saving and opening the document
DOC_PATH = Path.cwd() / "testing2" / "template.docx" 
word = win32.gencache.EnsureDispatch("Word.Application")
word.Visible = True

# Create a new document and ensure it is visible
doc = word.Documents.Add()
hwnd = win32gui.FindWindow("OpusApp", None)
win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
win32gui.SetForegroundWindow(hwnd)

# Page margin settings
doc.PageSetup.TopMargin    = cm_to_pt(1.7)
doc.PageSetup.BottomMargin = cm_to_pt(1.7)
doc.PageSetup.LeftMargin   = cm_to_pt(2.1)
doc.PageSetup.RightMargin  = cm_to_pt(1.7)

# Delete any default/old content
doc.Content.Delete()

# =================================================================================

# Starting the cursor in the beginning of the document
cursor = doc.Range(0, 0)
cursor.Collapse(c.wdCollapseEnd)

# Set Font, Size and Alignment 
cursor.Select()
word.Selection.Font.Name = "Times New Roman"
word.Selection.Font.Size = 15
word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter
word.Selection.ParagraphFormat.LineSpacingRule = c.wdLineSpaceSingle
word.Selection.Font.Bold = True       
word.Selection.Font.Italic = False

# Heading
word.Selection.TypeText(
    "VISVESVARAYA TECHNOLOGICAL UNIVERSITY\n"
    "“Jnana Sangama”, Belagavi – 590 018"
)
word.Selection.TypeParagraph()  
time.sleep(0.1)

# ---------------------------------------------------------------------------------

# VTU Logo
cursor = word.Selection.Range
cursor.Collapse(c.wdCollapseEnd)

# Insert a marker paragraph that will remain below the image
cursor.InsertAfter("\n")
cursor.Collapse(c.wdCollapseStart)
marker_range = cursor.Duplicate  # Save the location

# Insert the image inline and center it
image_path = str(Path.cwd() / "testing2" / "assets" / "VTU_Logo.png")
cursor.InsertParagraphAfter()
cursor.Collapse(c.wdCollapseEnd)
cursor.Select()
word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter

inline_shape = doc.InlineShapes.AddPicture(image_path, False, True, cursor)
inline_shape.LockAspectRatio = True
inline_shape.Width = cm_to_pt(4)

# Move cursor below image
cursor = inline_shape.Range.Duplicate
cursor.Collapse(c.wdCollapseEnd)
cursor.InsertParagraphAfter()
cursor.Collapse(c.wdCollapseEnd)
cursor.Select()

# ---------------------------------------------------------------------------------


# cursor.Select()
# word.Selection.TypeParagraph()

word.Selection.Font.Name = "Times New Roman"
word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter
word.Selection.ParagraphFormat.LineSpacingRule = c.wdLineSpaceSingle
word.Selection.Font.Bold = True       
word.Selection.Font.Italic = False

word.Selection.Font.Size = 10
word.Selection.TypeText(
    "A MINI PROJECT\n"
    "on\n"
)

time.sleep(0.1)
word.Selection.Font.Size = 16
word.Selection.TypeText(
    "\"  \"\n"   
)

time.sleep(0.1)
word.Selection.Font.Size = 10
word.Selection.Font.Bold = False
word.Selection.Font.Italic = True
word.Selection.TypeText(
    "Submitted in partial fulfillment of the requirements for the award of degree"   
)

time.sleep(0.1)
word.Selection.TypeParagraph()  


# ---------------------------------------------------------------------------------

# Insert a page break (New Page) 
cursor = word.Selection.Range
cursor.Collapse(c.wdCollapseEnd)
cursor.InsertBreak(c.wdPageBreak)
cursor.Select()
word.Selection.Range.GoTo()

# Border in Section 1
sec1 = doc.Sections(1)
borders = sec1.Borders
borders.DistanceFromTop = borders.DistanceFromBottom = 24
borders.DistanceFromLeft = borders.DistanceFromRight = 12

for side in (c.wdBorderTop, c.wdBorderLeft, c.wdBorderBottom, c.wdBorderRight):
    
    br = borders(side)
    br.LineStyle = c.wdLineStyleThinThickThinMedGap
    br.LineWidth = c.wdLineWidth300pt
    br.Color = c.wdColorAutomatic

# =================================================================================

# Saving the document
doc.SaveAs(str(DOC_PATH), FileFormat=c.wdFormatDocumentDefault)
print("✅ Created:", DOC_PATH.resolve())
