# Imports
from tkinter import *
import customtkinter as tk
from PIL import ImageTk

# Main class
class App(tk.CTk):
    def __init__(self):
        super().__init__()

        # Initialization settings
        self.windims = (900, 600)
        self.screen_width = self.winfo_screenwidth()
        self.screen_height = self.winfo_screenheight()
        x_coordinate = int((self.screen_width/2)-(self.windims[0]/2))
        y_coordinate = int((self.screen_height/2)-(self.windims[1]/2))-35

        self.geometry(f"{self.windims[0]}x{self.windims[1]}+{x_coordinate}+{y_coordinate}")
        self.resizable(False, False)
        self.title("Report Generator")

        icon_path = ImageTk.PhotoImage(file = r"v1\icon.png")
        self.wm_iconbitmap()
        self.iconphoto(False, icon_path)

        self.pages()
    
    def pages(self):
        # Format: (label_text, input_type, height)
        self.pages = [
            [("Name", "entry", 1), ("Age", "entry", 1), ("Bio", "text", 4)],
            [("Address", "text", 3)],
            [("Summary", "text", 5), ("Time", "entry", 1)],
            [("Signature", "entry", 1), ("Date", "entry", 1)]
        ]

        # Inputs
        self.current_page = 0
        self.user_inputs = [{} for _ in self.pages]

        self.title_label = tk.CTkLabel(self, text = "REPORT GENERATOR", font = ("Arial", 24, "bold"))
        self.title_label.pack(pady = 30)

        self.input_frame = tk.CTkFrame(self)
        self.input_frame.pack(pady = 40)

        self.entries = []  # List of (label, widget, type) tuples currently on the screen 

        # Buttons
        self.button_frame = tk.CTkFrame(self)
        self.button_frame.pack(side = "bottom", fill = "x", pady=30, padx=20)

        self.prev_button = tk.CTkButton(self.button_frame, text = "← Previous", command=self.go_previous)
        self.prev_button.pack(side = "left")

        self.next_button = tk.CTkButton(self.button_frame, text = "Next →", command=self.go_next)
        self.next_button.pack(side = "right")

        self.load_page()

    def load_page(self):
        # Clear current widgets
        for widget in self.input_frame.winfo_children():
            widget.destroy()
        self.entries.clear()

        # Current page info
        current_fields = self.pages[self.current_page]
        saved_data = self.user_inputs[self.current_page]

        for label_text, input_type, height in current_fields:
            label = tk.CTkLabel(self.input_frame, text = label_text, font = ("Arial", 16))
            label.pack(pady = (10, 2))

            common_fg_color = "#2A2D2E"  

            # Input widgets
            if input_type == "entry":
                widget = tk.CTkEntry(self.input_frame, width = 450, fg_color = common_fg_color)
                widget.pack(pady = (0, 10))
                if label_text in saved_data:
                    widget.insert(0, saved_data[label_text])

            elif input_type == "text":
                textbox_border = tk.CTkFrame(self.input_frame, fg_color = "#565b5e", corner_radius = 6)
                textbox_border.pack(pady = (0, 10), padx = 4)

                widget = tk.CTkTextbox(
                    textbox_border, width = 440, height = height * 30,
                    wrap = "word", fg_color = common_fg_color, corner_radius = 6
                )
                widget.pack(padx = 1.5, pady = 1.5)  

                if label_text in saved_data:
                    widget.insert("1.0", saved_data[label_text])
            else:
                continue  # Unknown input type

            self.entries.append((label_text, widget, input_type))

        # Change buttons according to pages
        self.prev_button.configure(state = "normal" if self.current_page > 0 else "disabled")
        if self.current_page == len(self.pages) - 1:
            self.next_button.configure(text = "Generate")
        else:
            self.next_button.configure(text = "Next →")

    # Save inputted records
    def save_current_inputs(self):
        page_data = {}
        for label_text, widget, input_type in self.entries:
            if input_type == "entry":
                page_data[label_text] = widget.get()
            elif input_type == "text":
                page_data[label_text] = widget.get("1.0", tk.END).strip()
        self.user_inputs[self.current_page] = page_data

    # last page
    def go_previous(self):
        self.save_current_inputs()
        if self.current_page > 0:
            self.current_page -= 1
            self.load_page()

    # next page
    def go_next(self):
        self.save_current_inputs()
        if self.current_page < len(self.pages) - 1:
            self.current_page += 1
            self.load_page()
        else:
            self.generate_report()

    # placeholder for report generator 
    def generate_report(self):
        print("Final Report Data:")
        for page_index, page in enumerate(self.user_inputs):
            print(f"Page {page_index + 1}:")
            for key, value in page.items():
                print(f"  {key}: {value}")


def main() -> None:
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
