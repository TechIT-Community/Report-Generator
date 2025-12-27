"""
Static content generation for the report (Title Page, Certificates, Acknowledgement, etc.).
Responsible for setting up the document structure, static text, placeholders, and formatting.
"""

import ctypes
import win32gui
import win32con
from win32com.client import constants as c
from pathlib import Path

from .utils import cm_to_pt
from .formatting import set_format, add_bookmark


# =================================================================================================
#                                      LAYOUT HELPERS
# =================================================================================================

def position_windows(word, doc):
    """
    Positions the Word window and the GUI application side by side.
    
    Layout:
    - [ GUI Application (Left 45%) ] [ Word Document (Right 55%) ]
    
    Sequence:
    1. Restore window (SW_SHOWNORMAL) if minimized.
    2. Wait for restore (polling).
    3. Set Position & Size using win32gui.SetWindowPos.
    4. Set Zoom (110%) and Scroll to middle.
    
    :param word: The Word Application object.
    :param doc: The active Document object.
    """
    screen_width = ctypes.windll.user32.GetSystemMetrics(0) # 1920 typ.
    screen_height = ctypes.windll.user32.GetSystemMetrics(1) # 1080 typ.

    half_width = screen_width // 2
    height = int(screen_height * 0.99)

    left = int(max(0, half_width - 0.107 * screen_width))
    width = int((half_width + 0.11 * screen_width))

    hwnd_word = win32gui.FindWindow("OpusApp", None) # Find the Word window
    
    if hwnd_word:
        import time
        
        # 1. Restore the window
        win32gui.ShowWindow(hwnd_word, win32con.SW_SHOWNORMAL)
        
        # 2. Polling loop to ensure window is actually restored
        for _ in range(20): # Wait up to 2 seconds
            is_minimized = win32gui.IsIconic(hwnd_word)
            if not is_minimized:
                break
            time.sleep(0.1)
            
        time.sleep(0.2) 

        # 3. Position the window
        if not win32gui.IsIconic(hwnd_word):
            try:
                win32gui.SetWindowPos( 
                    hwnd_word, None,
                    left, 0,
                    width, height,
                    win32con.SWP_NOZORDER
                )
            except Exception as e:
                print(f"⚠️ Failed to position Word window: {e}")
                
            # 4. Bring to foreground
            try:
                win32gui.SetForegroundWindow(hwnd_word)
            except Exception:
                pass
        else:
            print("⚠️ Word window failed to restore, skipping positioning.")

    # 5. Zoom and Scroll
    import time
    time.sleep(0.2) 

    try:
        if doc:
            doc.ActiveWindow.View.Zoom.Percentage = 110
        else:
            word.ActiveWindow.View.Zoom.Percentage = 110
    except Exception:
        pass # Ignore zoom errors

    window = word.ActiveWindow 
    if doc:
        window.ScrollIntoView(doc.Range(0, doc.Content.End // 2), True)


def make_borders(doc, word):
    """
    Applies a standard border to the first section of the document.
    
    :param doc: The Word Document object.
    :param word: The Word Application object.
    """
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


def delete_part2_content(doc):
    """
    Deletes all content after the Part1End bookmark (TOC, Chapters, References).
    
    Purpose:
    - This function allows the report to be regenerated multiple times in the same session.
    - It clears Part 2 (Dynamic Content) while keeping Part 1 (Static Content) intact.
    - Used when the user changes chapter count/content and presses "Done" again.

    Logic:
    1. Locate 'Part1End' bookmark (at end of Abstract).
    2. Delete everything from that point to end of document.
    3. Remove any extra sections created during previous generation.
    
    :param doc: The Word Document object.
    :return: True if deletion was successful, False otherwise.
    """
    # Part1End bookmark marks the boundary between static and dynamic content
    if not doc.Bookmarks.Exists("Part1End"):
        return False
    
    try:
        part1_end = doc.Bookmarks("Part1End").Range.End
        doc_end = doc.Content.End
        
        # Nothing to delete if Part1End is at the end
        if part1_end >= doc_end - 1:
            return True
        
        # Delete everything from Part1End to end of document
        delete_range = doc.Range(part1_end, doc_end)
        delete_range.Delete()
        
        # Remove extra sections created by Part 2 (keep only sections 1-2)
        while doc.Sections.Count > 2:
            last_section = doc.Sections(doc.Sections.Count)
            last_section.Range.Delete()
        
        return True
        
    except Exception:
        return False


def page_numbers(doc):
    """
    [DEPRECATED] This function is no longer used.
    Please use `page_numbers_dynamic` instead, which handles variable chapter counts correctly.
    
    Legacy config:
    - No numbers on Title/Certificate pages (Section 1-2).
    - Roman numerals for Section 3.
    """
    for idx, sec in enumerate(doc.Sections, start=1):
        sec.Range.InsertAfter("\r")
        if idx > 1:
            for hf_type in [c.wdHeaderFooterPrimary, c.wdHeaderFooterFirstPage]:
                sec.Footers(hf_type).LinkToPrevious = False
                sec.Headers(hf_type).LinkToPrevious = False

        # Sections 1 & 2: No numbering
        if idx == 1 or idx == 2:
            for hf_type in [c.wdHeaderFooterPrimary, c.wdHeaderFooterFirstPage]:
                sec.Footers(hf_type).Range.Text = ""
                sec.Headers(hf_type).Range.Text = ""
            continue

        # Section 3: Start numbering (usually distinct logic, here simplified to restart)
        if idx == 3:
            sec.PageSetup.DifferentFirstPageHeaderFooter = False
            footer = sec.Footers(c.wdHeaderFooterPrimary)
            pnums = footer.PageNumbers
            pnums.RestartNumberingAtSection = True
            pnums.StartingNumber = 1
            pnums.Add(c.wdAlignParagraphCenter, False)

        # Sections 4-8 (Chapters): Continue numbering
        if idx >= 4 and idx < 8:
            sec.PageSetup.DifferentFirstPageHeaderFooter = True
            pfooter = sec.Footers(c.wdHeaderFooterPrimary)
            ppnums = pfooter.PageNumbers
            ppnums.RestartNumberingAtSection = False
            ppnums.Add(c.wdAlignParagraphCenter, False)

            sec.Footers(c.wdHeaderFooterFirstPage).Range.Text = ""


# =================================================================================================
#                                   MAIN GENERATION LOGIC
# =================================================================================================

def generate_static_pages_part1(doc, word, base_dir: Path):
    """
    PART 1: Inserts static content blocks BEFORE the chapters.
    Sequence: Title -> Certificate -> Acknowledgement -> Abstract.
    
    NOTE: TOC, Chapters, and References are generated in Part 2.
    
    :param doc: The Word Document object.
    :param word: The Word Application object.
    :param base_dir: Base directory path for loading assets (images).
    """
    position_windows(word, doc)
    
    # Global cursor logic was used in original, here we use Selection mostly
    word.Selection.Range.Select()
    
    # ---------------------------------------------------------------------------------------------
    #                                     TITLE PAGE
    # ---------------------------------------------------------------------------------------------
    
    # Title formatting
    set_format(word.Selection, size=15, bold=True, align=c.wdAlignParagraphCenter, underline=c.wdUnderlineNone)

    position_windows(word, doc)
    word.Selection.TypeText(
        "VISVESVARAYA TECHNOLOGICAL UNIVERSITY\n"
        "“Jnana Sangama”, Belagavi – 590 018"
    )
    word.Selection.TypeParagraph()

    # -- VTU Logo Insertion --
    cursor = word.Selection.Range 
    cursor.Collapse(c.wdCollapseEnd) 
    word.Selection.TypeParagraph() 
    cursor.Collapse(c.wdCollapseStart) 
    
    image_path = str(base_dir / "assets" / "VTU_Logo.png")
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter

    inline_shape = doc.InlineShapes.AddPicture(image_path, False, True, cursor) 
    inline_shape.LockAspectRatio = True 
    inline_shape.Width = cm_to_pt(4) 

    cursor = inline_shape.Range.Duplicate 
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()

    # -- Project Title and Metadata --
    word.Selection.Font.Size = 11
    word.Selection.TypeText("A MINI PROJECT\vOn")
    word.Selection.TypeParagraph()
    
    set_format(word.Selection, size=15, bold=True, align=c.wdAlignParagraphCenter)
    add_bookmark(doc, word.Selection, "ProjectTitle", "___\n")

    set_format(word.Selection, size=11, bold=False, align=c.wdAlignParagraphCenter)
    word.Selection.Font.Italic = True
    word.Selection.TypeText("Submitted in partial fulfilment of the requirements for the award of degree")
    word.Selection.TypeParagraph()

    set_format(word.Selection, size=11, bold=False, align=c.wdAlignParagraphCenter)
    word.Selection.Font.Italic = False
    word.Selection.TypeText("Bachelor of Engineering\vIn\v")

    word.Selection.Font.Bold = True
    add_bookmark(doc, word.Selection, "Department", "___")
    word.Selection.TypeParagraph()    

    word.Selection.Font.Bold = False
    word.Selection.TypeText("Submitted by")
    word.Selection.TypeParagraph()    

    word.Selection.Font.Bold = True
    add_bookmark(doc, word.Selection, "NameAndUSN", "___\n")

    # -- Guidance Section (Guide & HOD) --
    word.Selection.Font.Bold = False
    word.Selection.TypeText("Under the guidance of\v")
    
    word.Selection.Font.Bold = True
    add_bookmark(doc, word.Selection, "GuideName", "___")
    word.Selection.TypeText("\v")
 
    word.Selection.Font.Bold = False
    add_bookmark(doc, word.Selection, "Designation", "___")
    word.Selection.TypeText("\v")

    # -- BNMIT Footer Logo --
    cursor = word.Selection.Range 
    cursor.Collapse(c.wdCollapseEnd) 
    word.Selection.TypeParagraph() 
    cursor.Collapse(c.wdCollapseStart)
    
    image_path = str(base_dir / "assets" / "BNMIT_Logo.png")
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter

    inline_shape = doc.InlineShapes.AddPicture(image_path, False, True, cursor) 
    inline_shape.LockAspectRatio = True 
    inline_shape.Width = cm_to_pt(5) 

    cursor = inline_shape.Range.Duplicate 
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()

    word.Selection.Font.Bold = True
    add_bookmark(doc, word.Selection, "Department_2", "___\n")
    
    if doc.Bookmarks.Exists("Department_2"):
         doc.Bookmarks("Department_2").Range.Case = c.wdUpperCase

    cursor = word.Selection.Range 
    cursor.Collapse(c.wdCollapseEnd) 
    
    # -- BNMIT Text Logo --
    image_path = str(base_dir / "assets" / "BNMIT_Text.png")
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter

    inline_shape = doc.InlineShapes.AddPicture(image_path, False, True, cursor) 
    inline_shape.LockAspectRatio = True 
    inline_shape.Width = cm_to_pt(15) 

    cursor = inline_shape.Range.Duplicate 
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()

    cursor.Collapse(c.wdCollapseEnd)
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1) 
    
    # Move to Next Page
    cursor.InsertBreak(c.wdPageBreak)
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
    
    # ---------------------------------------------------------------------------------------------
    #                                     CERTIFICATE PAGE
    # ---------------------------------------------------------------------------------------------

    cursor = word.Selection.Range 
    cursor.Collapse(c.wdCollapseEnd)
    
    # -- BNMIT Text Logo (Header) --
    image_path = str(base_dir / "assets" / "BNMIT_Text.png")
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

    # -- Department Header --
    placeholder = "___\n"
    word.Selection.TypeText(placeholder)
    bm_range = word.Selection.Range.Duplicate
    bm_start = bm_range.Start - len(placeholder)
    bm_range = doc.Range(bm_start, bm_start + len(placeholder))
    doc.Bookmarks.Add("Department_3", bm_range)
    bm_range.Case = c.wdUpperCase 

    # -- BNMIT Logo (Center) --
    cursor = word.Selection.Range 
    cursor.Collapse(c.wdCollapseEnd) 
    word.Selection.TypeParagraph()
    cursor.Collapse(c.wdCollapseStart)
    
    image_path = str(base_dir / "assets" / "BNMIT_Logo.png")
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

    # -- Certificate Body Text --
    word.Selection.Font.Name = "Calibri"                           
    word.Selection.Font.Size = 15                                          
    word.Selection.Font.Bold = True                                                
    word.Selection.Font.Italic = False                                       
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphCenter     
    word.Selection.ParagraphFormat.LineSpacingRule = c.wdLineSpace1pt5    
    word.Selection.Font.Underline = c.wdUnderlineSingle

    word.Selection.TypeText("CERTIFICATE")
    word.Selection.TypeParagraph()

    word.Selection.Font.Name = "Times New Roman"                            
    word.Selection.Font.Size = 12                                          
    word.Selection.Font.Bold = False                                                
    word.Selection.Font.Italic = False                                       
    word.Selection.ParagraphFormat.Alignment = c.wdAlignParagraphJustify     
    word.Selection.ParagraphFormat.LineSpacingRule = c.wdLineSpace1pt5    
    word.Selection.Font.Underline = c.wdUnderlineNone

    word.Selection.TypeText("This is to certify that the Mini project work entitled ")
    set_format(word.Selection, underline=c.wdUnderlineNone)
    
    set_format(word.Selection, bold=True)
    add_bookmark(doc, word.Selection, "ProjectTitle_2", "___")
    
    set_format(word.Selection, bold=False)
    word.Selection.TypeText(" is a bonafide work carried out by ")

    set_format(word.Selection, bold=True)
    add_bookmark(doc, word.Selection, "NameAndUSN_2", "___\n")
    
    set_format(word.Selection, bold=False)
    word.Selection.TypeText(" in partial fulfilment for the award of degree of ")

    set_format(word.Selection, bold=True)
    word.Selection.TypeText("Bachelor of Engineering")
    set_format(word.Selection, bold=False)
    word.Selection.TypeText(" in ")
    set_format(word.Selection, bold=True)
    add_bookmark(doc, word.Selection, "Department_4", "___") 
    
    set_format(word.Selection, bold=False)
    word.Selection.TypeText(" of the ")
    set_format(word.Selection, bold=True)
    word.Selection.TypeText("Visvesvaraya Technological University, Belagavi")
    set_format(word.Selection, bold=False)
    word.Selection.TypeText(" during the year ")
    set_format(word.Selection, bold=True)
    add_bookmark(doc, word.Selection, "Year", "___")
    
    set_format(word.Selection, bold=False)
    word.Selection.TypeText(". It is certified that all corrections/suggestions indicated for Internal Assessment have been incorporated in the report deposited in the departmental library. The project report has been approved as it satisfies the academic requirements in respect of Project work prescribed for the said Degree.")

    # -- Signature Table (Guide, HOD, Principal) --
    data = [
        ["___",     "___", "Dr. S Y Kulkarni"],
        ["___,",       "Professor and HOD,", "Additional Director"],
        ["___,",     "___,",      "and Principal,"],
        ["BNMIT, Bengaluru", "BNMIT, Bengaluru",   "BNMIT, Bengaluru"]
    ]
    bold_cells = [(0, 0), (0, 1), (0, 2)]
    
    cursor.Collapse(c.wdCollapseEnd)
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1) 
    cursor.Select()

    table = doc.Tables.Add(cursor, NumRows=len(data), NumColumns=max(len(r) for r in data))
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

    # Hide borders for signature table
    for border_id in [c.wdBorderTop, c.wdBorderBottom, c.wdBorderLeft, c.wdBorderRight, c.wdBorderHorizontal, c.wdBorderVertical]:
        border = table.Borders(border_id)
        border.LineStyle = c.wdLineStyleSingle
        border.Color = c.wdColorWhite

    cursor = table.Range.Duplicate
    cursor.Collapse(c.wdCollapseEnd)
    cursor.InsertParagraphAfter()
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()

    # -- Examiners Table (Header) --
    data = [["", "Name", "Signature with Date"]]
    bold_cells = [(0, 1), (0, 2)]
    cursor.Collapse(c.wdCollapseEnd)
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1) 
    cursor.Select()
    
    table = doc.Tables.Add(cursor, NumRows=len(data), NumColumns=max(len(r) for r in data))
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

    for border_id in [c.wdBorderTop, c.wdBorderBottom, c.wdBorderLeft, c.wdBorderRight, c.wdBorderHorizontal, c.wdBorderVertical]:
        border = table.Borders(border_id)
        border.LineStyle = c.wdLineStyleSingle
        border.Color = c.wdColorWhite

    cursor = table.Range.Duplicate
    cursor.Collapse(c.wdCollapseEnd)
    cursor.InsertParagraphAfter()
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()

    # -- Examiners Table (Rows) --
    data = [["Examiner 1:", "", ""], ["Examiner 2:", "", ""]]
    bold_cells = [(0, 0), (1, 0)]
    cursor.Collapse(c.wdCollapseEnd)
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1) 
    cursor.Select()
    
    table = doc.Tables.Add(cursor, NumRows=len(data), NumColumns=max(len(r) for r in data))
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
    
    for border_id in [c.wdBorderTop, c.wdBorderBottom, c.wdBorderLeft, c.wdBorderRight, c.wdBorderHorizontal, c.wdBorderVertical]:
        border = table.Borders(border_id)
        border.LineStyle = c.wdLineStyleSingle
        border.Color = c.wdColorWhite

    cursor = table.Range.Duplicate
    cursor.Collapse(c.wdCollapseEnd)
    cursor.InsertParagraphAfter()
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()

    # ---------------------------------------------------------------------------------------------
    #                                   ACKNOWLEDGEMENT PAGE
    # ---------------------------------------------------------------------------------------------
    
    cursor.Collapse(c.wdCollapseEnd)
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1) 
    cursor.InsertBreak(c.wdPageBreak) 
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()

    # -- Header --
    word.Selection.ParagraphFormat.LineSpacingRule = c.wdLineSpace1pt5
    set_format(word.Selection, size=14, bold=True, align=c.wdAlignParagraphCenter, underline=c.wdUnderlineNone)
    word.Selection.TypeText("ACKNOWLEDGEMENT")
    word.Selection.TypeParagraph()

    # -- Body Paragraphs --
    set_format(word.Selection, size=12, bold=False, align=c.wdAlignParagraphJustify)
    word.Selection.TypeText("I take this opportunity to express my heartfelt gratitude to all those who supported and guided me throughout the development of this project, ")
    set_format(word.Selection, bold=True)
    add_bookmark(doc, word.Selection, "ProjectTitle_Ack", "___") 
    set_format(word.Selection, bold=False)
    word.Selection.TypeText(". Their contributions and encouragement were invaluable to the successful completion of this endeavour.")
    word.Selection.TypeParagraph()
    word.Selection.TypeParagraph()

    word.Selection.TypeText("First and foremost, I would like to extend my sincere thanks to the Dean of our institution, Prof. Eishwar N Maanay, for providing the resources and a conducive environment to undertake this project. Their constant support and emphasis on innovation inspired me to push my boundaries.")
    word.Selection.TypeParagraph()
    word.Selection.TypeParagraph()

    word.Selection.TypeText("I am immensely grateful to our Head of the Department, ")
    set_format(word.Selection, bold=True)
    add_bookmark(doc, word.Selection, "HODName_Ack", "___")
    set_format(word.Selection, bold=False)
    word.Selection.TypeText(", ")
    add_bookmark(doc, word.Selection, "Department_9", "___")
    word.Selection.TypeText(" for their unwavering support and guidance. Their insights and suggestions played a crucial role in shaping the direction of this project. Their encouragement throughout the process has been a source of great motivation.")
    word.Selection.TypeParagraph()
    word.Selection.TypeParagraph()

    word.Selection.TypeText("A special note of appreciation goes to my Guide, ")
    set_format(word.Selection, bold=True)
    add_bookmark(doc, word.Selection, "GuideName_Ack", "___")
    set_format(word.Selection, bold=False)
    word.Selection.TypeText(", ")
    add_bookmark(doc, word.Selection, "Designation_Ack", "___")
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
   
    word.Selection.InsertBreak(c.wdPageBreak)
    word.Selection.MoveLeft(Unit=1, Count=1)
    word.Selection.Delete(Unit=1, Count=1)
    word.Selection.MoveRight(Unit=1, Count=1)

    # ---------------------------------------------------------------------------------------------
    #                                       ABSTRACT PAGE
    # ---------------------------------------------------------------------------------------------

    set_format(word.Selection, size=14, bold=True, align=c.wdAlignParagraphCenter, underline=c.wdUnderlineNone)
    word.Selection.TypeText("ABSTRACT")
    word.Selection.TypeParagraph()

    word.Selection.ParagraphFormat.LineSpacingRule = c.wdLineSpace1pt5    
    set_format(word.Selection, size=12, bold=False, align=c.wdAlignParagraphJustify)
    add_bookmark(doc, word.Selection, "Abstract", "___")

    cursor.Collapse(c.wdCollapseEnd)
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1) 
    cursor.InsertBreak(c.wdSectionBreakNextPage) 
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()
    
    # Mark end of Part 1 with a bookmark for Part 2 regeneration
    part1_end_range = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    doc.Bookmarks.Add("Part1End", part1_end_range)

    # PART 1 ENDS HERE. TOC, Chapters, and References are handled in Part 2.


# =================================================================================================
#                                   PART 2: DYNAMIC CHAPTERS
# =================================================================================================

def generate_static_pages_part2(doc, word, base_dir: Path, num_chapters: int):
    """
    PART 2: Generates dynamic sections based on user's chapter count.
    Sequence: TOC -> Chapters (1 to N) -> References.
    
    :param doc: The Word Document object.
    :param word: The Word Application object.
    :param base_dir: Base directory path for loading assets (images).
    :param num_chapters: The number of chapters to generate.
    """
    # ---------------------------------------------------------------------------------------------
    #                                     TABLE OF CONTENTS
    # ---------------------------------------------------------------------------------------------

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

    # -- Dynamic TOC Table Structure --
    data = [["S.No", "Title", "Page No"]]
    for i in range(1, num_chapters + 1):
        data.append([str(i), "___", "___"])
    data.append([str(num_chapters + 1), "References", "___"])
    
    bold_cells = [(0, 0), (0, 1), (0, 2)]
    
    cursor.Collapse(c.wdCollapseEnd)
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1) 
    cursor.Select()

    table = doc.Tables.Add(cursor, NumRows=len(data), NumColumns=max(len(r) for r in data))
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
    
    # -- Initialize Bookmarks using Table Cells --
    for i, row in enumerate(data):
        for j, cell_val in enumerate(row):
            cell = table.Cell(i + 1, j + 1)
            cell.Range.Text = cell_val
            if (i, j) in bold_cells:
                cell.Range.Font.Bold = True
            # Chapter Title placeholders (Column 2, Rows 1 to N)
            if j == 1 and 1 <= i <= num_chapters:
                placeholder = "___"
                cell.Range.Text = placeholder
                bm_start = cell.Range.Start
                bm_range = doc.Range(bm_start, bm_start + len(placeholder))
                doc.Bookmarks.Add(f"Chapter{i}Title", bm_range)
            # Chapter Page Number placeholders (Column 3, Rows 1 to N)
            if j == 2 and 1 <= i <= num_chapters:
                placeholder = "___"
                cell.Range.Text = placeholder
                bm_start = cell.Range.Start
                bm_range = doc.Range(bm_start, bm_start + len(placeholder))
                doc.Bookmarks.Add(f"Chapter{i}Page", bm_range)
            # References Page Number (Last Row, Column 3)
            if j == 2 and i == num_chapters + 1:
                placeholder = "___"
                cell.Range.Text = placeholder
                bm_start = cell.Range.Start
                bm_range = doc.Range(bm_start, bm_start + len(placeholder))
                doc.Bookmarks.Add("RefPage", bm_range)
                
    for border_id in [c.wdBorderTop, c.wdBorderBottom, c.wdBorderLeft, c.wdBorderRight, c.wdBorderHorizontal, c.wdBorderVertical]:
        border = table.Borders(border_id)
        border.LineStyle = c.wdLineStyleSingle
        border.Color = c.wdColorBlack

    cursor = table.Range.Duplicate
    cursor.Collapse(c.wdCollapseEnd)
    cursor.InsertParagraphAfter()
    cursor.Collapse(c.wdCollapseEnd)
    cursor.Select()

    # ---------------------------------------------------------------------------------------------
    #                                     CHAPTER CONTENT (Dynamic)
    # ---------------------------------------------------------------------------------------------

    for i in range(1, num_chapters + 1):
        cursor.Collapse(c.wdCollapseEnd)
        cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        cursor.InsertBreak(c.wdSectionBreakNextPage)
        cursor.Collapse(c.wdCollapseEnd)
        cursor.Select()

        word.Selection.Font.Name = "Times New Roman"
        set_format(word.Selection, size=16, bold=True, align=c.wdAlignParagraphCenter)

        for _ in range(9):
            word.Selection.TypeParagraph()
    
        # -- Chapter Title Placeholders --
        word.Selection.TypeText(f"Chapter {i}")
        word.Selection.TypeParagraph()
        placeholder = "___"
        word.Selection.TypeText(placeholder)
        bm_range = word.Selection.Range.Duplicate
        bm_start = bm_range.Start - len(placeholder)
        bm_range = doc.Range(bm_start, bm_start + len(placeholder))
        doc.Bookmarks.Add(f"Chapter{i}Title_2", bm_range)
        word.Selection.TypeParagraph()

        cursor.Collapse(c.wdCollapseEnd)
        cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        cursor.InsertBreak(c.wdPageBreak)
        cursor.Collapse(c.wdCollapseEnd)
        cursor.Select()

        # -- Chapter Title Repeat (Page 2) --
        placeholder = "___"
        word.Selection.TypeText(placeholder)
        bm_range = word.Selection.Range.Duplicate
        bm_start = bm_range.Start - len(placeholder)
        bm_range = doc.Range(bm_start, bm_start + len(placeholder))
        doc.Bookmarks.Add(f"Chapter{i}Title_3", bm_range)
        word.Selection.TypeParagraph()

        # -- Chapter Body Content --
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

    # ---------------------------------------------------------------------------------------------
    #                                     REFERENCES
    # ---------------------------------------------------------------------------------------------

    cursor.Collapse(c.wdCollapseEnd)
    cursor = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    cursor.InsertBreak(c.wdSectionBreakNextPage)
    cursor.Collapse(c.wdCollapseEnd)
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

    # ---------------------------------------------------------------------------------------------
    #                                  FINAL TOUCHES (Format & Numbers)
    # ---------------------------------------------------------------------------------------------

    make_borders(doc, word)
    page_numbers_dynamic(doc, num_chapters)


def page_numbers_dynamic(doc, num_chapters: int):
    """
    Configures page numbering logic for dynamic chapter count.
    
    Section Structure:
    - Section 1: Title + Certificate + Acknowledgement + Abstract (no page numbers)
    - Section 2: TOC (Roman numerals: i, ii, iii...)
    - Sections 3 to (2 + N): Chapters 1-N (Arabic: 1, 2, 3..., first page hidden)
    - Last Section: References (Arabic, continues from chapters)
    
    :param doc: The Word Document object.
    :param num_chapters: The number of chapters.
    """
    total_sections = doc.Sections.Count
    
    for idx, sec in enumerate(doc.Sections, start=1):
        # Insert paragraph to ensure section is properly separated
        sec.Range.InsertAfter("\r")
        
        # Break header/footer links so each section can have independent formatting
        if idx > 1:
            for hf_type in [c.wdHeaderFooterPrimary, c.wdHeaderFooterFirstPage]:
                sec.Footers(hf_type).LinkToPrevious = False
                sec.Headers(hf_type).LinkToPrevious = False

        # Section 1: Front matter (no page numbers)
        if idx == 1:
            for hf_type in [c.wdHeaderFooterPrimary, c.wdHeaderFooterFirstPage]:
                sec.Footers(hf_type).Range.Text = ""
                sec.Headers(hf_type).Range.Text = ""
            continue

        # Section 2: Table of Contents (Roman numerals starting at i)
        if idx == 2:
            sec.PageSetup.DifferentFirstPageHeaderFooter = False
            footer = sec.Footers(c.wdHeaderFooterPrimary)
            pnums = footer.PageNumbers
            pnums.RestartNumberingAtSection = True
            pnums.StartingNumber = 1
            pnums.NumberStyle = c.wdPageNumberStyleLowercaseRoman
            pnums.Add(c.wdAlignParagraphCenter, False)

        # Sections 3 to (2 + N): Chapter pages (Arabic, first page footer hidden)
        if 3 <= idx <= 2 + num_chapters:
            sec.PageSetup.DifferentFirstPageHeaderFooter = True
            pfooter = sec.Footers(c.wdHeaderFooterPrimary)
            ppnums = pfooter.PageNumbers
            
            # Use Arabic numerals for all chapters
            ppnums.NumberStyle = c.wdPageNumberStyleArabic
            
            # Chapter 1 restarts numbering at 1; others continue
            if idx == 3:
                ppnums.RestartNumberingAtSection = True
                ppnums.StartingNumber = 1
            else:
                ppnums.RestartNumberingAtSection = False
                
            ppnums.Add(c.wdAlignParagraphCenter, False)
            sec.Footers(c.wdHeaderFooterFirstPage).Range.Text = ""  # Hide first page footer

        # References: Last section, continues Arabic numbering from chapters
        if idx == total_sections and idx > 2 + num_chapters:
            sec.PageSetup.DifferentFirstPageHeaderFooter = False
            
            # Break link to prevent inheriting chapter's hidden first-page footer
            sec.Footers(c.wdHeaderFooterPrimary).LinkToPrevious = False
            sec.Headers(c.wdHeaderFooterPrimary).LinkToPrevious = False
            
            # Configure page numbers: Arabic, continue from chapters
            pfooter = sec.Footers(c.wdHeaderFooterPrimary)
            pnums = pfooter.PageNumbers
            pnums.NumberStyle = c.wdPageNumberStyleArabic
            pnums.RestartNumberingAtSection = False
            
            # Insert page number field in centered footer
            footer_range = pfooter.Range
            footer_range.Text = ""
            footer_range.ParagraphFormat.Alignment = c.wdAlignParagraphCenter
            footer_range.Fields.Add(footer_range, c.wdFieldPage)

