"""
Basic Graphical User Interface (GUI) for a Report Generator using CustomTKinter.

Key Features:
- Page-by-page input wizard (10 pages).
- Dynamic navigation (Next, Prev, Jump).
- Image uploads per chapter.
- Keyboard shortcuts for efficiency.
- Communication with the Word Automation Backend.

Dependencies: CustomTkinter, CTkMessagebox, PIL, backend.generator.
"""

from tkinter import *  # Standard Tkinter for basic GUI
import customtkinter as tk  # Modern UI
from CTkMessagebox import CTkMessagebox
from pathlib import Path  # Path handling
from tkinter import filedialog
import shutil

# Importing backend assumes running as module from package root
import app.backend.generator as docgen  

# =================================================================================================
#                                       CONFIGURATION
# =================================================================================================

# Adjusted to go up two levels: app/frontend/gui.py -> app/frontend -> app -> assets
BASE_DIR = Path(__file__).resolve().parent.parent 
ASSET_DIR = BASE_DIR / "assets"  # Directory for assets 


# =================================================================================================
#                                     MAIN APPLICATION CLASS
# =================================================================================================

class App(tk.CTk):
    """
    Main Application Window.
    Inherits from customtkinter.CTk to provide a modern, dark-themed UI.
    """
    
    def __init__(self, user_inputs):
        """
        Initializes the main window, layout, and event bindings.
        
        :param user_inputs: List of dictionaries containing pre-filled data (e.g., from Start Screen).
        """
        super().__init__()

        self.help_window = None 

        # --- Window Setup ---
        screen_w, screen_h = self.winfo_screenwidth(), self.winfo_screenheight()
        self.windims = (int(screen_w // 2 - 0.105 * screen_w), int(screen_h * 0.95))

        x = -(int(0.0057 * screen_w))
        y = int(((screen_h / 2) - (self.windims[1] / 2)) - (0.023 * screen_h))
        self.geometry(f"{self.windims[0]}x{self.windims[1]}+{x}+{y}")
        self.resizable(False, False)
        self.title("Report Generator")

        icon_path = str(ASSET_DIR / "icon.ico")
        self.iconbitmap(icon_path)

        # --- State Management ---
        self.uploaded_files = []
        self.user_inputs = user_inputs
        self.key_prefix_active = False
        self.floating_label_timer_id = None
        
        # --- Layout Initialization ---
        self.pages()
        self.user_inputs = user_inputs
        self.after(500, lambda: self.focus())
        docgen.initialize() # Initialize Word (Lazy Load)
        
        # Shortcut Label
        self.shortcut_label = tk.CTkLabel(self, text="F1: Keyboard shortcuts", font=("Arial", 12), text_color="gray")
        self.shortcut_label.place(relx=0.97, rely=0.03, anchor="ne")
        
        # --- Key Bindings ---
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

    # ---------------------------------------------------------------------------------------------
    #                                  UI FEEDBACK (FLASH LABEL)
    # ---------------------------------------------------------------------------------------------
    
    def flash_label(self, text, color="lightgreen", time = 1500):
        """Displays a temporary feedback message at the bottom of the window."""
        self.floating_label.configure(text=text, text_color=color)
        
        if self.floating_label_timer_id:
            self.after_cancel(self.floating_label_timer_id)

        self.floating_label_timer_id = self.after(time, lambda: self.floating_label.configure(text=""))
        
    def _show_next_right(self):
        """Visual wrapper for Next action."""
        if self.current_page < len(self.pages):
            self.flash_label(f"➡️ Next → Page {self.current_page + 1}: {self.page_titles[self.current_page]}")
            self.go_next()

    def _show_next_enter(self):
        """Visual wrapper for Enter key action."""
        if self.current_page < len(self.pages):
            self.flash_label(f"➡️ Next → Page {self.current_page + 1}: {self.page_titles[self.current_page]}")
            self.go_next()
        else:
            self.flash_label("✅ Done! Report saved successfully.", color="skyblue", time = 5000)
            self.save_entire_report()
            
    def _show_prev(self):
        """Visual wrapper for Previous action."""
        if self.current_page > 1:
            self.flash_label(f"⬅️ Back to Page {self.current_page - 1}: {self.page_titles[self.current_page - 2]}")
            self.go_previous()

    def _show_save(self):
        """Visual wrapper for Save action."""
        self.apply_page()
        self.flash_label("💾 Saved current page!")

    # ---------------------------------------------------------------------------------------------
    #                                   LIFECYCLE & HELPERS
    # ---------------------------------------------------------------------------------------------

    def on_close(self):
        """Cleanup handler when closing the window."""
        for file in self.uploaded_files:
            if file.exists() and file.name.startswith("Fig"):
                try:
                    file.unlink()
                except Exception as e:
                    print(f"⚠️ Couldn't delete {file.name}: {e}")
        self.destroy()
        
    def save_entire_report(self):
        """Calls the backend to finalize and save the Word document."""
        docgen.save_document()

    # ---------------------------------------------------------------------------------------------
    #                                     PAGE NAVIGATION
    # ---------------------------------------------------------------------------------------------

    def activate_page_jump_mode(self, event=None):
        self.key_prefix_active = True
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
        self.flash_label("🔚 Jumped to last page — press Done to finish.")

    # ---------------------------------------------------------------------------------------------
    #                                     HELP / SHORTCUTS
    # ---------------------------------------------------------------------------------------------

    def show_shortcuts_popup(self, event=None):
        """Displays a popup window with keyboard shortcuts."""
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
        
    # ---------------------------------------------------------------------------------------------
    #                                     PAGE CONTENT DEFINITIONS
    # ---------------------------------------------------------------------------------------------
        
    # ---------------------------------------------------------------------------------------------
    #                                     PAGE CONTENT DEFINITIONS
    # ---------------------------------------------------------------------------------------------
        
    def pages(self):
        """Defines the structure and fields for all 6 pages."""
        # Pages 1-4: Standard Info
        # Page 5: "Chapters" (Contains Tabs for Ch1-Ch5)
        # Page 6: References
        
        self.pages = [
            [("College", "entry", 1), ("Department", "entry", 1)],
            [("Project Title", "entry", 1), ("Name And USN", "text", 3), ("Guide Name", "entry", 1), ("Designation", "entry", 1)],
            [("Name USN", "text", 3), ("Sem", "entry", 1), ("Year", "entry", 1)],
            [("Abstract", "text", 5)],
            "CHAPTERS_TAB_INTERFACE", # Special Marker for Page 5
            [("References", "text", 5)]
        ]
        
        self.page_titles = [
            "College and Department",
            "Title Page",
            "Certificate Page",
            "Acknowledgement Page",
            "Chapters",
            "References"
        ]

        self.current_page = 1
        
        # --- TAB STATE ---
        self.chapter_tabs = []    # Stores tab dicts: {"name": str, "frame": CTkFrame, "entries": [], "data": {}}
        self.active_tab = None

        self.title_label = tk.CTkLabel(self, text="REPORT GENERATOR", font=("Arial", 24, "bold"))
        self.title_label.pack(pady=30)
        
        self.page_title_label = tk.CTkLabel(self, text="", font=("Arial", 18, "italic"))
        self.page_title_label.pack()

        self.input_frame = tk.CTkFrame(self, fg_color = "#1a1a1a")
        self.input_frame.pack(pady=40, fill="both", expand=True, padx=40) # Expanded for tabs
        
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

        self.load_page()

    # ---------------------------------------------------------------------------------------------
    #                                  PAGE LOADING & RENDERING
    # ---------------------------------------------------------------------------------------------

    def load_page(self):
        """Renders the current page's input fields."""
        # Cleanup general inputs
        for widget in self.input_frame.winfo_children():
            widget.destroy()
        self.entries.clear()
        
        # Cleanup old tabs references from memory if leaving page 5
        if self.current_page != 5:
            self.chapter_tabs = []
            self.active_tab = None

        self.page_title_label.configure(text=f"{self.current_page}: {self.page_titles[self.current_page - 1]}")
        self.page_selector.set(f"{self.current_page}. {self.page_titles[self.current_page - 1]}")

        # Check for Special Page 5 (Chapters)
        current_page_def = self.pages[self.current_page - 1]
        
        if current_page_def == "CHAPTERS_TAB_INTERFACE":
            self.render_chapter_interface()
            return

        # STANDARD PAGE RENDERING
        saved_data = self.user_inputs[self.current_page] if self.current_page < len(self.user_inputs) else {}

        for label_text, input_type, height in current_page_def:
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

        self.update_nav_buttons()

    def update_nav_buttons(self):
        self.prev_button.configure(state="normal" if self.current_page > 1 else "disabled")
        self.next_button.configure(text="Done" if self.current_page == len(self.pages) else "Next →")

    # ---------------------------------------------------------------------------------------------
    #                                  CUSTOM TAB MANAGER (For Page 5)
    # ---------------------------------------------------------------------------------------------

    def render_chapter_interface(self):
        """Builds the custom tab controller for Chapters 1-5."""
        
        # 1. Tab Navigation Bar
        self.tab_bar = tk.CTkFrame(self.input_frame, fg_color="transparent")
        self.tab_bar.pack(fill="x", pady=(0, 10))
        
        # 2. Content Container
        self.tab_content_container = tk.CTkFrame(self.input_frame, fg_color="transparent")
        self.tab_content_container.pack(fill="both", expand=True)

        # Initialize Tabs if empty
        if not self.chapter_tabs:
            for i in range(1, 6):
                self.create_chapter_tab(i)
        
        # Restore active tab or default to Ch1
        if not self.active_tab and self.chapter_tabs:
            self.set_active_tab(self.chapter_tabs[0])
        else:
            # Re-render bar if returning to page
            self.render_tab_buttons()
            if self.active_tab:
                 self.active_tab["frame"].pack(fill="both", expand=True)

    def create_chapter_tab(self, number):
        """Creates data structure and UI frame for a Chapter Tab."""
        tab = {
            "name": f"Chapter {number}",
            "id": number,
            "frame": tk.CTkFrame(self.tab_content_container, fg_color="transparent"),
            "entries": [], # Stores widget refs specific to this tab
            "data": {}     # Local data cache
        }
        
        # Populate UI inside the frame (Hidden by default)
        self.build_chapter_ui(tab)
        self.chapter_tabs.append(tab)

    def build_chapter_ui(self, tab):
        """Constructs the standard Chapter Title/Content/Upload widgets inside the tab frame."""
        frame = tab["frame"]
        
        # Pre-load data from GLOBAL storage if exists
        saved_data = self.user_inputs[self.current_page] if self.current_page < len(self.user_inputs) else {}
        
        # Fields Definition
        title_key = f"Chapter{tab['id']}Title"
        content_key = f"Chapter{tab['id']}Content"
        
        # 1. Title Input
        tk.CTkLabel(frame, text=f"{tab['name']} Title:", font=("Arial", 16)).pack(pady=(5, 2))
        title_entry = tk.CTkEntry(frame, width=450, fg_color="#2A2D2E")
        title_entry.pack(pady=(0, 10))
        if title_key in saved_data:
            title_entry.insert(0, saved_data[title_key])
            
        # 2. Content Input
        tk.CTkLabel(frame, text=f"{tab['name']} Content:", font=("Arial", 16)).pack(pady=(5, 2))
        border = tk.CTkFrame(frame, fg_color="#565b5e", corner_radius=6)
        border.pack(pady=(0, 10), padx=4)
        content_text = tk.CTkTextbox(border, width=440, height=180, wrap="word", fg_color="#2A2D2E", border_color="#565b5e")
        content_text.pack(padx=1.5, pady=1.5)
        if content_key in saved_data:
             content_text.insert("1.0", saved_data[content_key])

        # 3. Upload Button
        tk.CTkLabel(frame, text=f"Upload images for {tab['name']}:", font=("Arial", 14)).pack(pady=(10, 2))
        upload_btn = tk.CTkButton(frame, text="Upload Images", 
                                  command=lambda ch=tab['id']: self.browse_and_upload_images(ch))
        upload_btn.pack(pady=(5, 10))
        
        # Store refs for scraping later
        tab["entries"].append((title_key, title_entry, "entry"))
        tab["entries"].append((content_key, content_text, "text"))

    def set_active_tab(self, tab):
        """Switches the visible Chapter Tab."""
        if self.active_tab:
            self.active_tab["frame"].pack_forget()
            
        self.active_tab = tab
        self.active_tab["frame"].pack(fill="both", expand=True)
        self.render_tab_buttons()

    def render_tab_buttons(self):
        """Redraws the tab navigation bar."""
        for widget in self.tab_bar.winfo_children():
            widget.destroy()
            
        for tab in self.chapter_tabs:
            is_active = (tab is self.active_tab)
            btn = tk.CTkButton(
                self.tab_bar,
                text=tab["name"],
                width=100,
                fg_color="#1f538d" if is_active else "#333333",
                hover_color="#2b71ba" if is_active else "#444444",
                command=lambda t=tab: self.set_active_tab(t)
            )
            btn.pack(side="left", padx=5, fill="x", expand=True)


    # ---------------------------------------------------------------------------------------------
    #                                  DATA HANDLING & FLOW
    # ---------------------------------------------------------------------------------------------

    def save_current_inputs(self):
        """Scrapes current input widgets and stores them in self.user_inputs."""
        
        # CASE 1: CHAPTERS TAB INTERFACE (Page 5)
        if self.current_page == 5 and self.chapter_tabs:
            # We must scrape ALL tabs, not just the active one, because user might have typed in others
            combined_data = {}
            
            for tab in self.chapter_tabs:
                for label, widget, typ in tab["entries"]:
                    if typ == "entry":
                         combined_data[label] = widget.get()
                    elif typ == "text":
                         combined_data[label] = widget.get("1.0", tk.END).strip()
                         
            # Merge into the single Page 5 data slot
            # Note: We overwrite completely to ensure latest state is captured
            self.user_inputs[self.current_page] = combined_data
            return

        # CASE 2: STANDARD PAGE
        page_data = {}
        for label, widget, typ in self.entries:
            if typ == "entry":
                page_data[label] = widget.get()
            elif typ == "text":
                page_data[label] = widget.get("1.0", tk.END).strip()
        self.user_inputs[self.current_page] = page_data

    def go_previous(self):
        self.save_current_inputs()
        if self.current_page > 1:
            self.current_page -= 1
            self.load_page()

    def go_next(self):
        self.save_current_inputs()
        
        # When sending to backend, we always send the current page's data
        # For Page 5, this now contains aggregated data for ALL chapters
        docgen.replace_bookmarks(self.user_inputs[self.current_page])

        if self.current_page < len(self.pages):
            self.current_page += 1
            self.load_page()
        else:
            docgen.save_document()
            
    def apply_page(self):
        self.save_current_inputs()
        docgen.replace_bookmarks(self.user_inputs[self.current_page])

    def jump_to_page(self, selection):
        try:
            page_num = int(selection.split(".")[0])
            self.save_current_inputs()
            self.current_page = page_num
            self.load_page()
        except Exception as e:
            print(f"Page jump failed: {e}")

    def browse_and_upload_images(self, ch_num):
        """
        Opens a file dialog for image selection and copies them to assets.
        Auto-increments figure numbers (Fig X.Y).
        """
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
            self.uploaded_files.append(dest)


# =================================================================================================
#                                         ENTRY POINT
# =================================================================================================

def launch_gui(college, department):
    """
    Launches the main GUI loop.
    
    :param college: The selected college name.
    :param department: The selected department name.
    """
    # Initialize user_inputs with page 1 already filled
    user_inputs = [{},  # dummy for index 0, unused
                   {"College": college, "Department": department}]
    
    # Needs to be sparse list for 6 pages (Indices 0-6, so size 7)
    # 0=Dummy, 1-4=Info, 5=Chapters, 6=References
    user_inputs.extend({} for _ in range(7)) 

    tk.set_appearance_mode("dark")
    tk.set_default_color_theme("dark-blue")
    app = App(user_inputs=user_inputs)
    app.protocol("WM_DELETE_WINDOW", app.on_close)
    app.mainloop()
