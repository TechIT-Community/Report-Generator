"""
Basic Graphical User Interface (GUI) for a Report Generator using CustomTKinter.
1. Input Fields based on sections/pages
2. Press next to save input and updae the document
3. Press previous to go back and edit inputs
4. Press Done to save the document

Uses CustomTkinter for modern UI elements and PIL for image handling.
Uses a backend module `Document_Generator` for document handling. 
"""

# =================================================================================
# Imports

from tkinter import * # Standard Tkinter for basic GUI
import customtkinter as tk # Modern UI
from PIL import ImageTk # Image handling
import Document_Generator as docgen  # backend module

# =================================================================================

# Main Application Class
class App(tk.CTk):
    """
    This class represents the main application window for the Report Generator.
    It initializes the window, sets its dimensions, and manages the pages for user input.
    It allows users to navigate through different sections of the report, input data,
    and save their inputs. The application uses CustomTkinter for a modern look and feel.

    :param tk: CTk: Inherits from CustomTkinter's CTk class for modern UI elements.
    """
# ---------------------------------------------------------------------------------
    def __init__(self):
        """
        Initializes the main application window with specific dimensions and title.
        1. Initializes parent class `CTk` and sets up the window dimensions.
        2. Sets the window title and icon.
        3. calls :func:`self.pages()` to set up the input pages.
        """
        super().__init__()

        screen_w, screen_h = self.winfo_screenwidth(), self.winfo_screenheight() # Get screen dimensions (1536, 864)
        self.windims = (int(screen_w // 2 - 0.105*screen_w), int(screen_h * 0.95))  # Set to Left half of screen (Split Screen view)

        x = -(int(0.0057 * screen_w)) # Keep as 0 if outside screen  
        y = int(((screen_h / 2) - (self.windims[1] / 2)) - (0.023 * screen_h))
        self.geometry(f"{self.windims[0]}x{self.windims[1]}+{x}+{y}") # Dimensions and Position
        self.resizable(False, False) 
        self.title("Report Generator")

        icon_path = ImageTk.PhotoImage(file="app/v1/assets/icon.png")
        self.iconphoto(False, icon_path) # Set window icon

        self.pages() # Setup the page structure for user inputs
        docgen.insert_static_content()  # Setup document basics on app launch

# ---------------------------------------------------------------------------------

    def pages(self):
        """
        Modularly Defines the structure of the input pages
        """
    
        # All input fields, modularly changable
        # Each page has a list of tuples, each tuple being an input field with:
        # 1. Label text
        # 2. Input type (entry or text)
        # 3. Height of the field
        self.pages = [
            [("Project Title", "entry", 1)],
            [("Address", "text", 3)],
            [("Summary", "text", 5), ("Time", "entry", 1)],
            [("Signature", "entry", 1), ("Date", "entry", 1)]
        ]
        
        self.current_page = 0
        self.user_inputs = [{} for _ in self.pages] # List of dictionaries to store user inputs for each page
# _________________________________________________________________________________

        # GUI Elements
        self.title_label = tk.CTkLabel(self, text="REPORT GENERATOR", font=("Arial", 24, "bold"))
        self.title_label.pack(pady=30)

        self.input_frame = tk.CTkFrame(self) # Frame to hold input widgets
        self.input_frame.pack(pady=40)

        self.entries = [] # List of input widgets

        self.button_frame = tk.CTkFrame(self) # Frame to hold buttons
        self.button_frame.pack(side="bottom", fill="x", pady=30, padx=20)

        self.prev_button = tk.CTkButton(self.button_frame, text="← Previous", command=self.go_previous)
        self.prev_button.pack(side="left")

        self.next_button = tk.CTkButton(self.button_frame, text="Next →", command=self.go_next)
        self.next_button.pack(side="right")

        self.load_page() # Load any page (this call loads first page)

# ---------------------------------------------------------------------------------

    def load_page(self):
        """
        Loads the current page's input fields into the input frame 
        and dynamically changes the buttons and input widgets
        based on the current page index.
        """
        
        for widget in self.input_frame.winfo_children(): # Clear/destory all old widgets from previous page
            widget.destroy()
        self.entries.clear() # Empty the old input widget list for new page

        current_fields = self.pages[self.current_page] # List of fields for current page
        saved_data = self.user_inputs[self.current_page] # Currently saved data for current page

        # Creating the input field
        for label_text, input_type, height in current_fields: 
            
            label_text = label_text.replace(" ", "")
            
            label = tk.CTkLabel(self.input_frame, text=label_text + ":", font=("Arial", 16)) # Title
            label.pack(pady=(10, 2))

            fg_color = "#2A2D2E"

            if input_type == "entry": # Simple entry widget
                widget = tk.CTkEntry(self.input_frame, width=450, fg_color=fg_color)
                widget.pack(pady=(0, 10))
                if label_text in saved_data:
                    widget.insert(0, saved_data[label_text]) # Refill widget if already saved

            elif input_type == "text": # Textbox widget for multi-line input
                border = tk.CTkFrame(self.input_frame, fg_color="#565b5e", corner_radius=6)
                border.pack(pady=(0, 10), padx=4)

                widget = tk.CTkTextbox(border, width=440, height=height*30, wrap="word", fg_color=fg_color)
                widget.pack(padx=1.5, pady=1.5)
                if label_text in saved_data:
                    widget.insert("1.0", saved_data[label_text]) # Refill widget if already saved

            self.entries.append((label_text, widget, input_type)) # Add to list of fields

        # Change button states
        self.prev_button.configure(state="normal" if self.current_page > 0 else "disabled") 
        self.next_button.configure(text="Done" if self.current_page == len(self.pages) - 1 else "Next →")

# ---------------------------------------------------------------------------------

    def save_current_inputs(self):
        """
        Saves the current inputs from the input fields into the user_inputs dictionary.
        This method iterates through the entries of the current page and stores the data
        in the user_inputs list at the index of the current page.
        """
        page_data = {} # Dictionary to hold the current page's data
        for label, widget, type in self.entries:
            label = label.replace(" ", "") # remove spaces as bookmarks dont accept spaces
            if type == "entry":
                page_data[label] = widget.get() 
            elif type == "text":
                page_data[label] = widget.get("1.0", tk.END).strip()
        self.user_inputs[self.current_page] = page_data # Save the current page's data
        
        

# ---------------------------------------------------------------------------------

    def go_previous(self):
        """
        Saves the current inputs and navigates to the previous page.
        """
        self.save_current_inputs() # save current inputs before navigating
        if self.current_page > 0:
            self.current_page -= 1
            self.load_page() # Reload new page

# ---------------------------------------------------------------------------------

    def go_next(self):
        """
        Saves the current inputs and navigates to the next page.
        If the current page is the last page, it saves the document.
        """
        self.save_current_inputs() # save current inputs before navigating
        docgen.replace_bookmarks(self.user_inputs[self.current_page]) # Update the document's bookmarks/placeholders with current inputs

        if self.current_page < len(self.pages) - 1:
            self.current_page += 1
            self.load_page() # Reload new page
        else:
            docgen.save_document() # Save the document

# =================================================================================

#__Main__#
def main():
    app = App() # Create an instance of the App class
    app.mainloop() # Start the main event loop of the Application

# =================================================================================

# Entry point of script
if __name__ == "__main__":
    main()

# =================================================================================