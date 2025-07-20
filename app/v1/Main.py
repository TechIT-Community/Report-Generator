"""
Basic Graphical User Interface (GUI) for a Report Generator using CustomTKinter.
1. Input Fields based on sections/pages
2. Press next to save input and update the document
3. Press previous to go back and edit inputs
4. Press Done to save the document

Uses CustomTkinter for modern UI elements and PIL for image handling.
Uses a backend module `Document_Generator` for document handling. 
"""

# =================================================================================
# Imports

from tkinter import *  # Standard Tkinter for basic GUI
import customtkinter as tk  # Modern UI
from PIL import ImageTk  # Image handling
import Document_Generator as docgen  # backend module
from pathlib import Path  # Path handling

# =================================================================================

BASE_DIR = Path(__file__).resolve().parent  # Base directory of the application
ASSET_DIR = BASE_DIR / "assets"  # Directory for assets 

# Main Application Class
class App(tk.CTk):
    def __init__(self):
        super().__init__()

        screen_w, screen_h = self.winfo_screenwidth(), self.winfo_screenheight()
        self.windims = (int(screen_w // 2 - 0.105 * screen_w), int(screen_h * 0.95))

        x = -(int(0.0057 * screen_w))
        y = int(((screen_h / 2) - (self.windims[1] / 2)) - (0.023 * screen_h))
        self.geometry(f"{self.windims[0]}x{self.windims[1]}+{x}+{y}")
        self.resizable(False, False)
        self.title("Report Generator")

        icon_path = str(ASSET_DIR / "icon.ico")
        self.iconbitmap(icon_path)

        self.pages()
        self.after(500, lambda: self.focus())
        docgen.insert_static_content()


# ---------------------------------------------------------------------------------

    def pages(self):
        self.pages = [
            [("Project Title", "entry", 1), ("Name And USN", "text", 3), ("Guide Name", "entry", 1)],
            [("Name USN", "text", 3), ("Sem", "entry", 1), ("Year", "entry", 1)],
            [("Abstract", "text", 5)],
            [("Chapter 1 Title", "entry", 1), ("Chapter 1 Content", "text", 6)],
            [("Chapter 2 Title", "entry", 1), ("Chapter 2 Content", "text", 6)],
            [("Chapter 3 Title", "entry", 1), ("Chapter 3 Content", "text", 6)],
            [("Chapter 4 Title", "entry", 1), ("Chapter 4 Content", "text", 6)],
            [("Chapter 5 Title", "entry", 1), ("Chapter 5 Content", "text", 6)],
            [("References", "text", 5)]
        ]

        self.current_page = 0
        self.user_inputs = [{} for _ in self.pages]

        self.title_label = tk.CTkLabel(self, text="REPORT GENERATOR", font=("Arial", 24, "bold"))
        self.title_label.pack(pady=30)

        self.input_frame = tk.CTkFrame(self)
        self.input_frame.pack(pady=40)

        self.entries = []

        self.button_frame = tk.CTkFrame(self)
        self.button_frame.pack(side="bottom", fill="x", pady=30, padx=20)

        self.prev_button = tk.CTkButton(self.button_frame, text="← Previous", command=self.go_previous)
        self.prev_button.pack(side="left")

        self.next_button = tk.CTkButton(self.button_frame, text="Next →", command=self.go_next)
        self.next_button.pack(side="right")

        self.load_page()

# ---------------------------------------------------------------------------------

    def load_page(self):
        for widget in self.input_frame.winfo_children():
            widget.destroy()
        self.entries.clear()

        current_fields = self.pages[self.current_page]
        saved_data = self.user_inputs[self.current_page]

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

                widget = tk.CTkTextbox(border, width=440, height=height * 30, wrap="word", fg_color=fg_color)
                widget.pack(padx=1.5, pady=1.5)
                if label_key in saved_data:
                    widget.insert("1.0", saved_data[label_key])

            self.entries.append((label_key, widget, input_type))

        self.prev_button.configure(state="normal" if self.current_page > 0 else "disabled")
        self.next_button.configure(text="Done" if self.current_page == len(self.pages) - 1 else "Next →")

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
        if self.current_page > 0:
            self.current_page -= 1
            self.load_page()

# ---------------------------------------------------------------------------------

    def go_next(self):
        self.save_current_inputs()
        docgen.replace_bookmarks(self.user_inputs[self.current_page])

        if self.current_page < len(self.pages) - 1:
            self.current_page += 1
            self.load_page()
        else:
            docgen.save_document()

# =================================================================================

def main():
    app = App()
    app.mainloop()

# =================================================================================

if __name__ == "__main__":
    main()
