"""
make_test10_sections.py
-----------------------
Creates test10.docx in the folder where you run this script.

Prerequisites
-------------
• Windows with desktop Microsoft Word installed
• pip install pywin32
"""

import win32com.client as win32
from win32com.client import constants as c
from pathlib import Path
import win32gui
import win32con
import time

# ----------------------------------------------------------------------
# 1.  Where to save
# ----------------------------------------------------------------------
DOC_PATH = Path.cwd() / "testing2" / "test1.docx"

# ----------------------------------------------------------------------
# 2.  Start Word (hidden)
# ----------------------------------------------------------------------
word = win32.gencache.EnsureDispatch("Word.Application")
word.Visible = True
doc = word.Documents.Add()

time.sleep(0.1)
hwnd = win32gui.FindWindow("OpusApp", None)
win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
win32gui.SetForegroundWindow(hwnd)

# Wipe any default paragraph so we start truly at character 0
doc.Content.Delete()

# ----------------------------------------------------------------------
# 3.  Write 10 pages, inserting section breaks after page 3 and 7
# ----------------------------------------------------------------------
cursor = doc.Range(0, 0)
blank_pages = {4, 8, 11}                 # pages right after each section break
for page in range(1, 16):
    cursor.Collapse(c.wdCollapseEnd)

    # ----- TEXT GENERATION (only change) -----
    if page not in blank_pages:
        cursor.InsertAfter(f"Page {page}: filler text. " * 100)
        cursor.Select()
        word.Selection.Range.GoTo()
        time.sleep(0.1)
    # -----------------------------------------

    cursor.InsertParagraphAfter()
    time.sleep(0.1)

    if page in (3, 7, 10):
        cursor.Collapse(c.wdCollapseEnd)
        cursor.InsertBreak(c.wdSectionBreakNextPage)
        cursor.Select()
        word.Selection.Range.GoTo()
        time.sleep(0.1)
    elif page < 15:
        cursor.Collapse(c.wdCollapseEnd)
        cursor.InsertBreak(c.wdPageBreak)
        cursor.Select()
        word.Selection.Range.GoTo()
        time.sleep(0.1)
# ----------------------------------------------------------------------
# 4.  Configure page-numbering rules per section
#     • Section 1 (index 1): no numbers
#     • Section 2 (index 2): restart at 1
#     • Section 3 (index 3): restart at 1
# ----------------------------------------------------------------------
for idx, sec in enumerate(doc.Sections, start=1):
    for hf_type in [c.wdHeaderFooterPrimary, c.wdHeaderFooterFirstPage]:
        sec.Footers(hf_type).LinkToPrevious = False

    if idx == 1:
        for hf_type in [c.wdHeaderFooterPrimary, c.wdHeaderFooterFirstPage]:
            sec.Footers(hf_type).Range.Text = ""
        continue

    if idx == 2:
        sec.Range.Select()                      
        word.Selection.Range.GoTo()
        time.sleep(0.1)

        sec.PageSetup.DifferentFirstPageHeaderFooter = False
        footer = sec.Footers(c.wdHeaderFooterPrimary)
        pnums = footer.PageNumbers
        pnums.RestartNumberingAtSection = True
        pnums.StartingNumber = 1
        pnums.Add(c.wdAlignParagraphCenter, False)

    if idx >= 3:
        sec.Range.Select()                      
        word.Selection.Range.GoTo()
        time.sleep(0.1)

        sec.PageSetup.DifferentFirstPageHeaderFooter = True
        pfooter = sec.Footers(c.wdHeaderFooterPrimary)
        ppnums = pfooter.PageNumbers
        ppnums.RestartNumberingAtSection = False
        ppnums.Add(c.wdAlignParagraphCenter, False)
        sec.Footers(c.wdHeaderFooterFirstPage).Range.Text = ""

# ----------------------------------------------------------------------
# 5. Border
# ----------------------------------------------------------------------
sec1 = doc.Sections(1)
borders = sec1.Borders
borders.DistanceFromTop = borders.DistanceFromBottom = 24
borders.DistanceFromLeft = borders.DistanceFromRight = 12

sec1.Range.Select()
word.Selection.Range.GoTo()
time.sleep(0.1)

for side in (c.wdBorderTop, c.wdBorderLeft, c.wdBorderBottom, c.wdBorderRight):
    
    br = borders(side)
    br.LineStyle = c.wdLineStyleThinThickThinMedGap
    br.LineWidth = c.wdLineWidth300pt
    br.Color = c.wdColorAutomatic

# ----------------------------------------------------------------------
# 5.5 Insert image and text on page 5
# ----------------------------------------------------------------------

# Adjust the page number here
target_page = 5

# 1. Find range near end of that page
go_range = doc.GoTo(What=c.wdGoToPage, Name=str(target_page))
cursor = go_range
cursor.Collapse(c.wdCollapseEnd)
cursor.Select()
word.Selection.Range.GoTo()
cursor.InsertAfter(f"Page {page}: filler text. " * 20)

# 2. Insert a paragraph before image
cursor.InsertParagraphAfter()
cursor.Collapse(c.wdCollapseEnd)


# 3. Insert the image
image_path = str(Path.cwd() / "testing2" / "sample_image.png")  # Replace with your image filename
inline_shape = doc.InlineShapes.AddPicture(image_path, False, True, cursor)
cursor.Select()
word.Selection.Range.GoTo()

# 4. Convert to floating shape with "Top and Bottom" wrapping
shape = inline_shape.ConvertToShape()
shape.WrapFormat.Type = c.wdWrapTopBottom
shape.Select()
# word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter  # optional: center it
time.sleep(0.1)

# 5. Insert text after the image
shape.Anchor.Collapse(c.wdCollapseEnd)
shape.Anchor.InsertAfter("\nText after the image on page 5.")
cursor.Select()
word.Selection.Range.GoTo()


# ----------------------------------------------------------------------
# 6.  Save & quit
# ----------------------------------------------------------------------

doc.SaveAs(str(DOC_PATH), FileFormat=c.wdFormatDocumentDefault)
# doc.Close(False)
# word.Quit()

print("✅ Created:", DOC_PATH.resolve())
