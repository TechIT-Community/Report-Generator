"""
Image insertion logic for the report generator.
Handles the dynamic discovery, resizing, and smart placement of images within chapters.
"""

from win32com.client import constants as c
from pathlib import Path
from PIL import Image
import re


# =================================================================================================
#                                  IMAGE INSERTION CONTROLLER
# =================================================================================================

def insert_images_in_chapter(doc, chapter_num: int, start_range, asset_dir: Path):
    """
    Scans the asset directory for images belonging to the specified chapter 
    (naming pattern: "Fig {chapter_num}.{index}.png") and inserts them after the content.
    
    Implements Smart Placement:
    - Calculates image dimensions before insertion.
    - Checks available vertical space on the current page.
    - Inserts a Page Break if the image + caption would overflow, preventing cut-off.
    
    :param doc: The Word Document object.
    :param chapter_num: Integer chapter number (1-5).
    :param start_range: The Range object representing the end of the chapter's text content.
    :param asset_dir: Directory to search for images.
    """
    
    # -------------------------- Image Discovery --------------------------
    
    def extract_figure_index(p):
        """Helper to sort images by their numeric index (Fig 1.1, 1.2, 1.10)."""
        match = re.search(rf"Fig {chapter_num}\.(\d+)", p.stem)
        if match:
            return float(match.group(1))
        return float('inf')

    image_files = sorted(
        asset_dir.glob(f"Fig {chapter_num}.*"),
        key=extract_figure_index
    )

    if not image_files:
        return

    # -------------------------- Range Calculation --------------------------
    
    # Define start of insertion range (immediately after text content)
    chapter_end = start_range.End

    # Define end of chapter (boundary) by checking next chapter title or document end
    # This prevents images from spilling into the next chapter's territory
    next_title = f"Chapter{chapter_num + 1}Title_2"
    if doc.Bookmarks.Exists(next_title):
        chapter_limit = doc.Bookmarks(next_title).Range.Start
    else:
        chapter_limit = doc.Content.End

    # Check for existing figure captions to avoid overlapping or duplicate insertion
    safe_start = min(chapter_end, chapter_limit)
    safe_end = max(chapter_end, chapter_limit)
    if safe_end > doc.Content.End:
        safe_end = doc.Content.End

    scan_range = doc.Range(safe_start, safe_end)
    existing_text = scan_range.Text

    # Prepare insertion cursor
    insert_range = doc.Range(chapter_end, chapter_end)
    insert_range.Collapse(c.wdCollapseStart)

    # FIX: Check for existing figures to append AFTER them (preserve order)
    # Search for the highest existing figure index in the text scan range
    existing_indices = []
    # Regex to find "Fig X.Y" where X is chapter_num
    matches = re.finditer(rf"Fig {chapter_num}\.(\d+)", existing_text)
    for m in matches:
        try:
            existing_indices.append(int(m.group(1)))
        except ValueError:
            pass
            
    if existing_indices:
        last_index = max(existing_indices)
        last_label = f"Fig {chapter_num}.{last_index}"
        
        # Locate this label in the document to move cursor after it
        search_rng = scan_range.Duplicate
        # Find the text "Fig X.Y"
        if search_rng.Find.Execute(FindText=last_label, Forward=True, Wrap=0): # 0 = wdFindStop
            # FOUND: search_rng now covers "Fig X.Y"
            # Move to end of this range
            search_rng.Collapse(c.wdCollapseEnd)
            
            # The caption is usually followed by a paragraph mark (inserted by InsertParagraphAfter)
            # Try to move past potential paragraph mark to start next insertion cleanly
            search_rng.MoveEnd(c.wdParagraph, 1) 
            search_rng.Collapse(c.wdCollapseEnd)
            
            insert_range = search_rng.Duplicate


    # -------------------------- Insertion Loop --------------------------

    for img in image_files:
        fig_index = img.stem.split('.')[-1]
        fig_label = f"Fig {chapter_num}.{fig_index}"

        # Skip if this specific figure label already exists in the zone
        if fig_label in existing_text:
            continue

        # --- Smart Placement Logic ---
        
        # 1. Calc target dimensions using PIL (without inserting yet)
        max_width_pt = 450 
        target_height_pt = 0
        try:
             with Image.open(str(img.resolve())) as pil_img:
                 w_px, h_px = pil_img.size
                 aspect = h_px / w_px
                 
                 # Convert px to pt (Approximate: 1 px = 0.75 pt at 96 DPI)
                 natural_width_pt = w_px * 0.75
                 
                 # Effective width uses natural size unless it exceeds page printable width
                 effective_width_pt = min(natural_width_pt, max_width_pt)
                 target_height_pt = effective_width_pt * aspect 
        except Exception:
             # Fallback if image reading fails
             target_height_pt = 200 # Arbitrary default

        # 2. Check available space on page
        try:
            wdVerticalPositionRelativeToPage = 6 # Constant
            current_vertical_pos = insert_range.Information(wdVerticalPositionRelativeToPage)
            
            # Get Page Height and Limit
            page_height = doc.PageSetup.PageHeight
            bottom_margin = doc.PageSetup.BottomMargin
            limit = page_height - bottom_margin
            
            available_space = limit - current_vertical_pos
            caption_buffer = 60 # Points reserved for caption text + spacing
            
            # 3. Decide on Page Break
            if (current_vertical_pos + target_height_pt + caption_buffer) > limit:
                # Not enough space, insert page break to move image to fresh page
                insert_range.InsertBreak(c.wdPageBreak)
                insert_range.Collapse(c.wdCollapseEnd)
                
        except Exception as e:
            print(f"⚠️ Calculation error: {e}. Letting Word decide placement.")
        
        # --- Physical Insertion ---
        
        img_range = insert_range.Duplicate
        img_shape = img_range.InlineShapes.AddPicture(str(img.resolve()), LinkToFile=False, SaveWithDocument=True)
        
        # Center the image
        img_shape.Range.ParagraphFormat.Alignment = c.wdAlignParagraphCenter
        img_shape.Range.ParagraphFormat.KeepWithNext = True # Keep image with its caption
        
        # --- Caption Insertion ---
        
        caption_range = img_shape.Range.Duplicate
        caption_range.Collapse(c.wdCollapseEnd)
        caption_range.InsertParagraphAfter()
        caption_range.Collapse(c.wdCollapseEnd)
        
        caption_range.Text = fig_label
        caption_range.Font.Name = "Times New Roman"
        caption_range.Font.Size = 12
        caption_range.Font.Bold = False
        caption_range.ParagraphFormat.Alignment = c.wdAlignParagraphCenter
        caption_range.ParagraphFormat.SpaceAfter = 12 
        
        caption_range.InsertParagraphAfter()
        
        # --- Advance Cursor ---
        insert_range = caption_range.Duplicate
        insert_range.Collapse(c.wdCollapseEnd)
