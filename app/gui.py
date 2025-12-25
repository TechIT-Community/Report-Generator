"""
Basic Graphical User Interface (GUI) for a Report Generator using CustomTKinter.
1. Input Fields based on sections/pages
2. Press next to save input and update the document
3. Press previous to go back and edit inputs
4. Press Done to save the document

Uses CustomTkinter for modern UI elements and PIL for image handling.
Uses a backend module `Document_Generator` for document handling. 
"""

from tkinter import *  # Standard Tkinter for basic GUI
import customtkinter as tk  # Modern UI
from CTkMessagebox import CTkMessagebox
from pathlib import Path  # Path handling
from tkinter import filedialog
import shutil

import Document_Generator as docgen  # backend module

# =================================================================================

BASE_DIR = Path(__file__).resolve().parent  # Base directory of the application
ASSET_DIR = BASE_DIR / "assets"  # Directory for assets 

# =================================================================================

class App(tk.CTk):
    def __init__(self, user_inputs):
        super().__init__()

        self.help_window = None 

        screen_w, screen_h = self.winfo_screenwidth(), self.winfo_screenheight()
        self.windims = (int(screen_w // 2 - 0.105 * screen_w), int(screen_h * 0.95))

        x = -(int(0.0057 * screen_w))
        y = int(((screen_h / 2) - (self.windims[1] / 2)) - (0.023 * screen_h))
        self.geometry(f"{self.windims[0]}x{self.windims[1]}+{x}+{y}")
        self.resizable(False, False)
        self.title("Report Generator")

        icon_path = str(ASSET_DIR / "icon.ico")
        self.iconbitmap(icon_path)

        self.uploaded_files = []

        self.user_inputs = user_inputs
        self.key_prefix_active = False
        self.floating_label_timer_id = None
        
        self.pages()
        self.user_inputs = user_inputs
        self.after(500, lambda: self.focus())
        docgen.insert_static_content()
        
        # ========== KEY BINDINGS ==========
        self.bind_all("<Control-Return>", lambda e: self._show_next_enter())  # Ctrl + Enter = Next
        self.bind_all("<Control-Right>", lambda e: self._show_next_right())  # Ctrl + → = Next
        self.bind_all("<Control-Left>", lambda e: self._show_prev())  # Ctrl + ← = Previous

        self.bind_all("<Control-s>", lambda e: self._show_save())
        self.bind_all("<Control-Shift-S>", lambda e: self.save_entire_report())
        self.bind_all("<Control-q>", self.jump_to_last_with_prompt)
        self.bind_all("<Escape>", self.jump_to_last_with_prompt)
        self.bind_all("<F1>", self.show_shortcuts_popup)

        # Page jump prefix mode
        self.bind_all("<Control-k>", self.activate_page_jump_mode)
        for i in range(1, 10):
            self.bind_all(str(i), lambda e, i=i: self.page_jump_prefix(i))
        self.bind_all("0", lambda e: self.page_jump_prefix(10))

# ---------------------------------------------------------------------------------
    def flash_label(self, text, color="lightgreen", time = 1500):
        self.floating_label.configure(text=text, text_color=color)
        
        if self.floating_label_timer_id:
            self.after_cancel(self.floating_label_timer_id)

        self.floating_label_timer_id = self.after(time, lambda: self.floating_label.configure(text=""))
        
    def _show_next_right(self):
        if self.current_page < len(self.pages):
            self.flash_label(f"🔄 Next → Page {self.current_page + 1}: {self.page_titles[self.current_page]}")
            self.go_next()

    def _show_next_enter(self):
        if self.current_page < len(self.pages):
            self.flash_label(f"🔄 Next → Page {self.current_page + 1}: {self.page_titles[self.current_page]}")
            self.go_next()
        else:
            self.flash_label("✅ Done! Report saved successfully.", color="skyblue", time = 5000)
            self.save_entire_report()
            

    def _show_prev(self):
        if self.current_page > 1:
            self.flash_label(f"⬅️ Back to Page {self.current_page - 1}: {self.page_titles[self.current_page - 2]}")
            self.go_previous()

    def _show_save(self):
        self.apply_page()
        self.flash_label("💾 Saved current page!")

# ---------------------------------------------------------------------------------

    def on_close(self):
        for file in self.uploaded_files:
            if file.exists() and file.name.startswith("Fig"):
                try:
                    file.unlink()
                except Exception as e:
                    print(f"⚠️ Couldn't delete {file.name}: {e}")
        self.destroy()
        
    def save_entire_report(self):
        docgen.save_document()
        #CTkMessagebox(title="Saved", message="Entire report saved successfully.", icon="check")

# ---------------------------------------------------------------------------------

    def activate_page_jump_mode(self, event=None):
        self.key_prefix_active = True
        #CTkMessagebox(title="Page Jump Mode", message="Press 1–0 to jump to page", icon="info")
        self.flash_label("⌨️ Page jump mode: Press 1–0")


    def jump_to_page_by_index(self, index, event=None):
        self.jump_to_page(f"{index}. {self.page_titles[index - 1]}")

    def page_jump_prefix(self, num):
        if self.key_prefix_active:
            self.key_prefix_active = False
            self.jump_to_page_by_index(num)
            self.flash_label(f"✅ Jumped to Page {num}: {self.page_titles[num - 1]}")

    def jump_to_last_with_prompt(self, event=None):
        self.save_current_inputs()
        self.current_page = len(self.pages)
        self.load_page()
        #CTkMessagebox(title="Ready to Submit", message="You're on the last page. Press 'Done' to save your report.", icon="info")
        self.flash_label("🔚 Jumped to last page — press Done to finish.")

# ---------------------------------------------------------------------------------

    def show_shortcuts_popup(self, event=None):
        if self.help_window and self.help_window.winfo_exists():
            self.help_window.destroy()
            self.help_window = None
            return

        self.help_window = tk.CTkToplevel(self)
        self.help_window.title("Shortcut Help")
        self.help_window.geometry("420x280")
        self.help_window.resizable(False, False)
        self.help_window.attributes("-topmost", True)

        heading = tk.CTkLabel(
            self.help_window,
            text="Keyboard Shortcuts\n",
            font=("Arial", 18, "bold"),
            text_color="skyblue",
        )
        heading.pack(pady=(15, 5))

        shortcuts_text = (
            "• F1: Show/Hide this help\n\n"
            "• Ctrl + Enter / Ctrl + →: Next\n"
            "• Ctrl + ←: Previous\n"
            "• Ctrl + S: Save current page\n"
            "• Ctrl + Shift + S: Save entire report\n"
            "• Ctrl + Q or Esc: Jump to last page and prompt\n\n"
            "• Ctrl + K, then 1–9 or 0: Jump to pages 1–10\n"
        )

        label = tk.CTkLabel(
            self.help_window,
            text=shortcuts_text,
            font=("Arial", 14),
            justify="left",
            wraplength=400
        )
        label.pack(padx=20, pady=(0, 20))
        
# ---------------------------------------------------------------------------------
        
    def pages(self):
        self.pages = [
            [("College", "entry", 1), ("Department", "entry", 1)],
            [("Project Title", "entry", 1), ("Name And USN", "text", 3), ("Guide Name", "entry", 1), ("Designation", "entry", 1)],
            [("Name USN", "text", 3), ("Sem", "entry", 1), ("Year", "entry", 1)],
            [("Abstract", "text", 5)],
            [("Chapter 1 Title", "entry", 1), ("Chapter 1 Content", "text", 6)],
            [("Chapter 2 Title", "entry", 1), ("Chapter 2 Content", "text", 6)],
            [("Chapter 3 Title", "entry", 1), ("Chapter 3 Content", "text", 6)],
            [("Chapter 4 Title", "entry", 1), ("Chapter 4 Content", "text", 6)],
            [("Chapter 5 Title", "entry", 1), ("Chapter 5 Content", "text", 6)],
            [("References", "text", 5)]
        ]
        
        self.page_titles = [
            "College and Department",
            "Title Page",
            "Certificate Page",
            "Acknowledgement Page",
            "Chapter 1",
            "Chapter 2",
            "Chapter 3",
            "Chapter 4",
            "Chapter 5",
            "References"
        ]

        self.current_page = 1

        self.title_label = tk.CTkLabel(self, text="REPORT GENERATOR", font=("Arial", 24, "bold"))
        self.title_label.pack(pady=30)
        
        self.page_title_label = tk.CTkLabel(self, text="", font=("Arial", 18, "italic"))
        self.page_title_label.pack()

        self.input_frame = tk.CTkFrame(self, fg_color = "#1a1a1a")
        self.input_frame.pack(pady=40)
        
        self.save_button = tk.CTkButton(self, text="💾 Save", command=self.apply_page)
        self.save_button.pack(pady=(10, 0))

        self.entries = []

        self.button_frame = tk.CTkFrame(self, fg_color = "#1a1a1a")
        self.button_frame.pack(side="bottom", fill="x", pady=30, padx=20)

        self.prev_button = tk.CTkButton(self.button_frame, text="← Previous", command=self._show_prev)
        self.prev_button.pack(side="left")

        self.next_button = tk.CTkButton(self.button_frame, text="Next →", command=self._show_next_enter)
        self.next_button.pack(side="right")
        
        self.page_selector = tk.CTkOptionMenu(
            self.button_frame,
            values=[f"{i+1}. {title}" for i, title in enumerate(self.page_titles)],
            command=self.jump_to_page
        )
        self.page_selector.pack(pady=5)

        self.floating_label = tk.CTkLabel(self, text="", font=("Arial", 14), text_color="lightgreen")
        self.floating_label.pack(side="bottom", pady=(5, 0))
        #self.floating_label.after(300, lambda: self.floating_label.configure(text=""))

        self.load_page()

# ---------------------------------------------------------------------------------

    def load_page(self):
        for widget in self.input_frame.winfo_children():
            widget.destroy()
        self.entries.clear()

        self.page_title_label.configure(text=f"Page {self.current_page}: {self.page_titles[self.current_page - 1]}")
        self.page_selector.set(f"{self.current_page}. {self.page_titles[self.current_page - 1]}")

        current_fields = self.pages[self.current_page - 1]
        saved_data = self.user_inputs[self.current_page] if self.current_page < len(self.user_inputs) else {}

        for label_text, input_type, height in current_fields:
            label_key = label_text.replace(" ", "")
            label = tk.CTkLabel(self.input_frame, text=label_text + ":", font=("Arial", 16))
            label.pack(pady=(10, 2))

            fg_color = "#2A2D2E"

            if input_type == "entry":
                widget = tk.CTkEntry(self.input_frame, width=450, fg_color=fg_color)
                widget.pack(pady=(0, 10))
                if label_key in saved_data:
                    widget.insert(0, saved_data[label_key])
            elif input_type == "text":
                border = tk.CTkFrame(self.input_frame, fg_color="#565b5e", corner_radius=6)
                border.pack(pady=(0, 10), padx=4)

                widget = tk.CTkTextbox(border, width=440, height=height * 30, wrap="word", fg_color=fg_color, border_color = "#565b5e")
                widget.pack(padx=1.5, pady=1.5)
                if label_key in saved_data:
                    widget.insert("1.0", saved_data[label_key])

            self.entries.append((label_key, widget, input_type))

            # If this is a chapter content field, allow image upload
            if label_key.startswith("Chapter") and "Content" in label_key and 4 <= self.current_page <= 9:
                chapter_number = self.current_page - 4  # Pages 4 to 8 → Chapters 1 to 5

                image_upload_frame = tk.CTkFrame(self.input_frame, fg_color="#1a1a1a")
                image_upload_frame.pack(pady=(0, 10))

                upload_label = tk.CTkLabel(image_upload_frame, text=f"Upload images for Chapter {chapter_number}:", font=("Arial", 14))
                upload_label.pack()

                def upload_images(ch_num=chapter_number):
                    files = filedialog.askopenfilenames(
                        title="Select image(s)",
                        filetypes=[("Image Files", "*.png *.jpg *.jpeg *.bmp *.gif")],
                    )
                    if not files:
                        return

                    existing = list(ASSET_DIR.glob(f"Fig {ch_num}.*"))
                    next_idx = 1 + max([
                        int(p.stem.split('.')[1]) for p in existing if p.stem.startswith(f"Fig {ch_num}.") and '.' in p.stem
                    ] + [0])

                    for i, path in enumerate(files, start=next_idx):
                        ext = Path(path).suffix.lower()
                        dest = ASSET_DIR / f"Fig {ch_num}.{i}{ext}"
                        shutil.copy(path, dest)
                        self.flash_label(f"📸 Uploaded: {dest.name}", time=2000)
                        #CTkMessagebox(title="Upload Successful", message=f"Saved as: {dest.name}", icon="check")
                        self.uploaded_files.append(dest)

                upload_btn = tk.CTkButton(image_upload_frame, text="Upload Images", command=upload_images)
                upload_btn.pack(pady=(5, 0))

        self.prev_button.configure(state="normal" if self.current_page > 1 else "disabled")
        self.next_button.configure(text="Done" if self.current_page == len(self.pages) else "Next →")

# ---------------------------------------------------------------------------------

    def save_current_inputs(self):
        page_data = {}
        for label, widget, typ in self.entries:
            if typ == "entry":
                page_data[label] = widget.get()
            elif typ == "text":
                page_data[label] = widget.get("1.0", tk.END).strip()
        self.user_inputs[self.current_page] = page_data

# ---------------------------------------------------------------------------------

    def go_previous(self):
        self.save_current_inputs()
        if self.current_page > 1:
            self.current_page -= 1
            self.load_page()

# ---------------------------------------------------------------------------------

    def go_next(self):
        self.save_current_inputs()
        docgen.replace_bookmarks(self.user_inputs[self.current_page])

        if self.current_page < len(self.pages):
            self.current_page += 1
            self.load_page()
        else:
            docgen.save_document()

# ---------------------------------------------------------------------------------
            
    def apply_page(self):
        self.save_current_inputs()
        docgen.replace_bookmarks(self.user_inputs[self.current_page])
        #CTkMessagebox(title="Saved", message="Changes applied to document.", icon="check")
    
# ---------------------------------------------------------------------------------

    def jump_to_page(self, selection):
        try:
            page_num = int(selection.split(".")[0])
            self.save_current_inputs()
            self.current_page = page_num
            self.load_page()
        except Exception as e:
            print(f"Page jump failed: {e}")

# =================================================================================

def launch_gui(college, department):
    # Initialize user_inputs with page 1 already filled
    user_inputs = [{},  # dummy for index 0, unused
                   {"College": college, "Department": department}]
    user_inputs.extend({} for _ in range(9))  # Total 10 pages (1-indexed)

    tk.set_appearance_mode("dark")
    tk.set_default_color_theme("dark-blue")
    app = App(user_inputs=user_inputs)
    app.protocol("WM_DELETE_WINDOW", app.on_close)
    app.mainloop()
