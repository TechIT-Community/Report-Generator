"""
Image insertion logic for the report generator.
"""
from win32com.client import constants as c
from pathlib import Path
from PIL import Image
import re

def insert_images_in_chapter(doc, chapter_num: int, start_range, asset_dir: Path):
    """
    Scans the asset directory for images belonging to the specified chapter 
    and inserts them after the content.
    """
    def extract_figure_index(p):
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

    # Step 1: Define start of insertion range
    chapter_end = start_range.End

    # Step 2: Define end of chapter by checking next chapter title or document end
    next_title = f"Chapter{chapter_num + 1}Title_2"
    if doc.Bookmarks.Exists(next_title):
        chapter_limit = doc.Bookmarks(next_title).Range.Start
    else:
        chapter_limit = doc.Content.End

    # Step 3: Define range to check for existing figure captions to avoid dupes
    safe_start = min(chapter_end, chapter_limit)
    safe_end = max(chapter_end, chapter_limit)
    if safe_end > doc.Content.End:
        safe_end = doc.Content.End

    scan_range = doc.Range(safe_start, safe_end)
    existing_text = scan_range.Text

    # Step 4: Begin inserting images using an advancing range
    insert_range = doc.Range(chapter_end, chapter_end)
    insert_range.Collapse(c.wdCollapseStart)

    for img in image_files:
        fig_index = img.stem.split('.')[-1]
        fig_label = f"Fig {chapter_num}.{fig_index}"

        if fig_label in existing_text:
            continue  # Already inserted

        # --- Smart Placement Logic ---
        # 1. Calc target dimensions
        max_width_pt = 450 
        target_height_pt = 0
        try:
             with Image.open(str(img.resolve())) as pil_img:
                 w_px, h_px = pil_img.size
                 aspect = h_px / w_px
                 
                 # Convert px to pt (Approximate: 1 px = 0.75 pt at 96 DPI)
                 natural_width_pt = w_px * 0.75
                 
                 # Effective width uses natural size unless it exceeds page width
                 effective_width_pt = min(natural_width_pt, max_width_pt)
                 target_height_pt = effective_width_pt * aspect 
        except Exception:
             # Fallback if image reading fails
             target_height_pt = 200 # Arbitrary default

        # 2. Check available space
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
            if (current_vertical_pos + target_height_pt + caption_buffer) > limit:
                # Not enough space, force page break
                insert_range.InsertBreak(c.wdPageBreak)
                insert_range.Collapse(c.wdCollapseEnd)
                
        except Exception as e:
            print(f"⚠️ Calculation error: {e}. Letting Word decide placement.")
        
        # Step 2: Insert image
        img_range = insert_range.Duplicate
        img_shape = img_range.InlineShapes.AddPicture(str(img.resolve()), LinkToFile=False, SaveWithDocument=True)
        
        # Center the image
        img_shape.Range.ParagraphFormat.Alignment = c.wdAlignParagraphCenter
        img_shape.Range.ParagraphFormat.KeepWithNext = True 
        
        # Step 3: Insert Caption
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
        
        # Step 4: Advance safely
        insert_range = caption_range.Duplicate
        insert_range.Collapse(c.wdCollapseEnd)
