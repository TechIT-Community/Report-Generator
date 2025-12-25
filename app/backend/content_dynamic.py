"""
Dynamic content replacement (bookmarks) and updates for the report.
Handles user data injection, image insertion, and final page number updates.
"""

import re
from win32com.client import constants as c
from pathlib import Path
from .images import insert_images_in_chapter


# =================================================================================================
#                                  BOOKMARK REPLACEMENT LOGIC
# =================================================================================================

def replace_bookmarks(doc, word, data_dict: dict, asset_dir: Path):
    """
    Replaces bookmarks in the Word document with values from a dictionary.
    Also inserts images after Chapter{i}Content bookmarks if matching files are found.
    
    1.  Transforms input data:
        -   Maps Departments to Short Forms (Dept. of CSE)
        -   Maps Departments to HOD Names
        -   Formats multiline names for certificates
    2.  Iterates through all document bookmarks.
    3.  Replaces matched bookmarks with text.
    4.  Triggers Image Insertion logic for Chapter Content.
    5.  Updates Headers and Footers with Project Title and Year.
    
    :param doc: The Word Document object.
    :param word: The Word Application object.
    :param data_dict: Dictionary of user inputs.
    :param asset_dir: Directory containing assets (images).
    """
    
    # -------------------------- Data Transformation --------------------------
    transformed_data = {}
    
    dept_short_forms = {
        "COMPUTER SCIENCE AND ENGINEERING": "Dept. of CSE",
        "ELECTRONICS AND COMMUNICATION ENGINEERING": "Dept. of ECE",
        "INFORMATION SCIENCE AND ENGINEERING": "Dept. of ISE",
        "MECHANICAL ENGINEERING": "Dept. of ME",
        "CIVIL ENGINEERING": "Dept. of CE",
        "ELECTRONICS AND INSTRUMENTATION ENGINEERING": "Dept. of EIE",
        "ARTIFICIAL INTELLIGENCE AND MACHINE LEARNING": "Dept. of AIML",
        "ELECTRICAL AND ELECTRONICS ENGINEERING": "Dept. of EEE"
    }
    
    hod_titles = {
        "COMPUTER SCIENCE AND ENGINEERING": "Dr. Chayadevi M.L",
        "ELECTRONICS AND COMMUNICATION ENGINEERING": "Dr. P. A. Vijaya", 
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
        # HOD full name → for Department_5 (Certificate)
        hod_value = hod_titles.get(department_value, department_value)
        transformed_data["Department_5"] = hod_value
        
        # Department_8 is used in "Department of [Department_8]". Should be full name or just branch.
        transformed_data["Department_8"] = department_value 

        # Short form dept → for Department_6 and Department_7
        short_form = dept_short_forms.get(department_value, department_value)
        transformed_data["Department_6"] = short_form
        transformed_data["Department_7"] = short_form
        transformed_data["Department_9"] = department_value # Changed to Full Name for Acknowledgement
        
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

    # Also carry over other keys from data_dict directly
    for key, value in data_dict.items():
        if key != "Department":  # Already handled separately
            if key == "NameAndUSN":
                # Special handling for Certificate Page usage
                # If NameAndUSN has multiline input, replace newlines with commas for inline certificate
                inline_names = value.replace("\n", ", ")
                transformed_data["NameAndUSN_2"] = inline_names
            
            transformed_data[key] = value
            
    all_bm_names = [bm.Name for bm in doc.Bookmarks]  # Get all bookmark names in the document

    # These bookmarks should have a newline after the inserted value
    # NOTE: GuideName and Designation removed from here to prevent layout breaks (handled in static)
    newline_bookmark_names = {
        "ProjectTitle", "NameAndUSN", 
        "Department_2", "Department_3",
        "Chapter1Title", "Chapter2Title", "Chapter3Title", "Chapter4Title", "Chapter5Title",
        "Chapter1Content", "Chapter2Content", "Chapter3Content", "Chapter4Content", "Chapter5Content"
    }

    rebookmarks = []  # To store bookmarks that need to be re-added after replacement

    # -------------------------- Replacement Loop --------------------------
    # Uses transformed_data to ensure derived keys are covered
    
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
            insert_text = value + ("\n" if add_newline else "") 
            
            bm_range.Text = insert_text
            
            new_range = doc.Range(bm_start, bm_start + len(insert_text))
            rebookmarks.append((name, new_range))
            
            new_range.Select()
            word.ActiveWindow.ScrollIntoView(word.Selection.Range, True)
            
            # --- Handle images (ChapterContent logic) ---
            chapter_match = re.match(r"Chapter(\d)Content", name)
            if chapter_match:
                chapter_num = int(chapter_match.group(1))
                insert_images_in_chapter(doc, chapter_num, new_range, asset_dir)

    # --- Re-add bookmarks ---
    for name, rng in rebookmarks:
        try:
            doc.Bookmarks.Add(name, rng)
        except:
            print(f"⚠️ Could not re-add bookmark: {name}")

    # -------------------------- Header / Footer Updates --------------------------
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


# =================================================================================================
#                                   PAGE NUMBER UPDATES (TOC/Index)
# =================================================================================================

def update_index_page_numbers(doc):
    """
    Updates the Table of Contents (TOC) page numbers by looking up the actual
    page numbers of the Chapter Titles and References.
    
    Uses `wdActiveEndAdjustedPageNumber` to handle section restarts correctly.
    """
    # Attempt to use wdActiveEndAdjustedPageNumber (4) for restart-aware numbering
    wdActiveEndAdjustedPageNumber = getattr(c, 'wdActiveEndAdjustedPageNumber', 4)

    # 1. Update Chapter 1-5 Page Numbers
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
            bm_range.Text = str(page_number) 

            # Re-bookmark the range
            new_range = doc.Range(bm_start, bm_start + len(str(page_number)))
            try:
                doc.Bookmarks.Add(page_bm, new_range)
            except:
                print(f"⚠️ Could not re-add bookmark: {page_bm}")
                
    # 2. Update Reference Page Number
    if doc.Bookmarks.Exists("References") and doc.Bookmarks.Exists("RefPage"):
        ref_range = doc.Bookmarks("References").Range
        ref_page = ref_range.Information(wdActiveEndAdjustedPageNumber) 

        bm_range = doc.Bookmarks("RefPage").Range
        bm_start = bm_range.Start
        bm_range.Text = str(ref_page)

        # Re-bookmark the range
        new_range = doc.Range(bm_start, bm_start + len(str(ref_page)))
        try:
            doc.Bookmarks.Add("RefPage", new_range)
        except:
            print(f"⚠️ Could not re-add bookmark: RefPage")
