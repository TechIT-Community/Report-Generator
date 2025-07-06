import win32com.client as win32
from win32com.client import constants as c
from pathlib import Path
import win32gui
import win32con
import time
import ctypes

# Globals
word = win32.gencache.EnsureDispatch("Word.Application")
word.Visible = True
DOC_PATH = Path.cwd() / "testing2" / "template.docx"
doc = word.Documents.Add()

# Setup Word window
hwnd = win32gui.FindWindow("OpusApp", None)
win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
win32gui.SetForegroundWindow(hwnd)

# Conversion helper
cm_to_pt = lambda cm: cm * 28.3464566929133858

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

# Position Word and GUI nicely side by side
def position_windows():
    screen_width = ctypes.windll.user32.GetSystemMetrics(0)
    screen_height = ctypes.windll.user32.GetSystemMetrics(1)

    half_width = screen_width // 2
    height = int(screen_height * 0.98)

    hwnd_word = win32gui.FindWindow("OpusApp", None)
    if hwnd_word:
        win32gui.ShowWindow(hwnd_word, win32con.SW_RESTORE)
        win32gui.SetWindowPos(
            hwnd_word, None,
            half_width, 0,
            half_width, height,
            win32con.SWP_NOZORDER
        )

# ------------------ Initial Static Content & Bookmarks ------------------

def insert_static_content():
    global cursor
    cursor.Select()
    word.Selection.Font.Name = "Times New Roman"
    word.Selection.Font.Size = 15
    word.Selection.Font.Bold = True
    word.Selection.Font.Italic = False
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter
    word.Selection.ParagraphFormat.LineSpacingRule = c.wdLineSpaceSingle

    # Heading
    word.Selection.TypeText(
        "VISVESVARAYA TECHNOLOGICAL UNIVERSITY\n"
        "“Jnana Sangama”, Belagavi – 590 018"
    )
    word.Selection.TypeParagraph()
    time.sleep(0.1)

    # Image
    cursor = word.Selection.Range
    cursor.Collapse(c.wdCollapseEnd)
    cursor.InsertAfter("\n")
    cursor.Collapse(c.wdCollapseStart)
    marker_range = cursor.Duplicate

    image_path = str(Path.cwd() / "testing2" / "assets" / "VTU_Logo.png")
    cursor.InsertParagraphAfter()
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter

    inline_shape = doc.InlineShapes.AddPicture(image_path, False, True, cursor)
    inline_shape.LockAspectRatio = True
    inline_shape.Width = cm_to_pt(4)

    cursor = inline_shape.Range.Duplicate
    cursor.Collapse(c.wdCollapseEnd)
    cursor.InsertParagraphAfter()
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()

    # Sub-heading
    word.Selection.Font.Size = 10
    word.Selection.TypeText("A MINI PROJECT\non\n")
    time.sleep(0.1)

    # Title placeholder with bookmark
    word.Selection.TypeText("___\n")
    title_range = word.Selection.Range.Duplicate
    title_range.MoveStart(Unit=c.wdCharacter, Count=-4)
    doc.Bookmarks.Add("projectTitle", title_range)

    time.sleep(0.1)
    word.Selection.Font.Size = 10
    word.Selection.Font.Bold = False
    word.Selection.Font.Italic = True
    word.Selection.TypeText("Submitted in partial fulfillment of the requirements for the award of degree")
    word.Selection.TypeParagraph()

    position_windows()  # Call to arrange Word window properly

# ------------------ Replace Bookmarks With GUI Input ------------------

def replace_bookmarks(data_dict: dict):
    print("Bookmarks in document:", [bm.Name for bm in doc.Bookmarks])
    for key, value in data_dict.items():
        if doc.Bookmarks.Exists(key):
            bm_range = doc.Bookmarks(key).Range
            bm_start = bm_range.Start
            bm_range.Text = value + "\n"
            new_range = doc.Range(bm_start, bm_start + len(value) + 1)
            doc.Bookmarks.Add(key, new_range)
            print(f"✔ Replaced bookmark '{key}' with '{value}'")
        else:
            print(f"Bookmark '{key}' not found in document. Skipping...")
    cursor = doc.Range()
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()

# ------------------ Save ------------------

def save_document():
    doc.SaveAs(str(DOC_PATH), FileFormat=c.wdFormatDocumentDefault)
    print("✅ Saved:", DOC_PATH.resolve())
