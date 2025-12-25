# gui.py

from tkinter import *
import customtkinter as tk
from PIL import ImageTk
import windocgen as docgen  # Import backend module

class App(tk.CTk):
    def __init__(self):
        super().__init__()

        screen_w, screen_h = self.winfo_screenwidth(), self.winfo_screenheight()
        self.windims = (screen_w // 2, int(screen_h * 0.95))  # Left half of screen

        x = 0  # Start at left edge
        y = int((screen_h / 2) - (self.windims[1] / 2))
        self.geometry(f"{self.windims[0]}x{self.windims[1]}+{x}+{y}")
        self.resizable(False, False)
        self.title("Report Generator")

        icon_path = ImageTk.PhotoImage(file="testing3/assets/icon.png")
        self.iconphoto(False, icon_path)

        self.pages()
        docgen.insert_static_content()  # Setup doc on app launch

    def pages(self):
        self.pages = [
            [("projectTitle", "entry", 1)],
            [("Address", "text", 3)],
            [("Summary", "text", 5), ("Time", "entry", 1)],
            [("Signature", "entry", 1), ("Date", "entry", 1)]
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

    def load_page(self):
        for widget in self.input_frame.winfo_children():
            widget.destroy()
        self.entries.clear()

        current_fields = self.pages[self.current_page]
        saved_data = self.user_inputs[self.current_page]

        for label_text, input_type, height in current_fields:
            label = tk.CTkLabel(self.input_frame, text=label_text + ":", font=("Arial", 16))
            label.pack(pady=(10, 2))

            fg_color = "#2A2D2E"

            if input_type == "entry":
                widget = tk.CTkEntry(self.input_frame, width=450, fg_color=fg_color)
                widget.pack(pady=(0, 10))
                if label_text in saved_data:
                    widget.insert(0, saved_data[label_text])

            elif input_type == "text":
                border = tk.CTkFrame(self.input_frame, fg_color="#565b5e", corner_radius=6)
                border.pack(pady=(0, 10), padx=4)

                widget = tk.CTkTextbox(border, width=440, height=height*30, wrap="word", fg_color=fg_color)
                widget.pack(padx=1.5, pady=1.5)
                if label_text in saved_data:
                    widget.insert("1.0", saved_data[label_text])

            self.entries.append((label_text, widget, input_type))

        self.prev_button.configure(state="normal" if self.current_page > 0 else "disabled")
        self.next_button.configure(text="Done" if self.current_page == len(self.pages) - 1 else "Next →")

    def save_current_inputs(self):
        page_data = {}
        for label, widget, typ in self.entries:
            if typ == "entry":
                page_data[label] = widget.get()
            elif typ == "text":
                page_data[label] = widget.get("1.0", tk.END).strip()
        self.user_inputs[self.current_page] = page_data

    def go_previous(self):
        self.save_current_inputs()
        if self.current_page > 0:
            self.current_page -= 1
            self.load_page()

    def go_next(self):
        self.save_current_inputs()
        docgen.replace_bookmarks(self.user_inputs[self.current_page])
        print(self.user_inputs[self.current_page], "\n\n")

        if self.current_page < len(self.pages) - 1:
            self.current_page += 1
            self.load_page()
        else:
            docgen.save_document()

def main():
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
