"""
Entry point for the Report Generator Application.
Displays a Start Screen for initial configuration (College, Department) before launching the main GUI.

Architecture:
1.  StartScreen (CTk): Gets user configuration.
2.  Main Loop orchestrator:
    - Runs StartScreen.
    - If successful, destroys StartScreen.
    - Imports and launching the main `gui` module.
    - This separation ensures clean window lifecycle management.
"""

from tkinter import *  # Standard Tkinter for basic GUI
import customtkinter as tk  # Modern UI
from CTkMessagebox import CTkMessagebox
from pathlib import Path  # Path handling
from PIL import Image  # Fix: Needed for CTkImage
import sys

# Add project root to sys.path to resolve 'app' package
project_root = Path(__file__).resolve().parent.parent.parent
if str(project_root) not in sys.path:
    sys.path.append(str(project_root))

# =================================================================================================
#                                       CONFIGURATION
# =================================================================================================

# Adjusted to go up two levels: app/frontend/main.py -> app/frontend -> app -> assets
BASE_DIR = Path(__file__).resolve().parent.parent 
ASSET_DIR = BASE_DIR / "assets"  # Directory for assets 

# =================================================================================================
#                                       START SCREEN
# =================================================================================================

class StartScreen(tk.CTk):
    """
    Initial configuration window.
    Allows the user to select their College and Department.
    """
    
    def __init__(self):
        super().__init__()

        # --- Window Configuration ---
        self.title("Start Report Generator")
        self.geometry("700x450")
        self.resizable(False, False)
        
        # Center Window
        self.update_idletasks()
        w = 700
        h = 450
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()

        x = int((screen_w - w) / 2 + 0.05 * screen_w)  
        y = int((screen_h - h) / 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

        icon_path = str(ASSET_DIR / "icon.ico")
        self.iconbitmap(icon_path)

        # --- UI Elements ---
        logo_path = ASSET_DIR / "icon.png"
        logo_image = Image.open(logo_path)
        self.logo = tk.CTkImage(light_image=logo_image, dark_image=logo_image, size=(120, 120))
        self.logo_label = tk.CTkLabel(self, image=self.logo, text="")
        self.logo_label.pack(pady=(30, 10))
        
        self.title_label = tk.CTkLabel(self, text="REPORT GENERATOR", font=("Arial", 24, "bold"))
        self.title_label.pack(pady=(0, 20))

        # Dropdowns
        self.college_var = tk.StringVar(value="Select College")
        self.college_menu = tk.CTkOptionMenu(
            self, values=["BNMIT", "more coming soon"], variable=self.college_var
        )
        self.college_menu.pack(pady=10)

        self.dept_var = tk.StringVar(value="Select Department")
        self.dept_menu = tk.CTkOptionMenu(
            self, values=[
                "COMPUTER SCIENCE AND ENGINEERING",
                "ELECTRONICS AND COMMUNICATION ENGINEERING",
                "INFORMATION SCIENCE AND ENGINEERING",
                "MECHANICAL ENGINEERING",
                "CIVIL ENGINEERING",
                "ELECTRONICS AND INSTRUMENTATION ENGINEERING",
                "ARTIFICIAL INTELLIGENCE AND MACHINE LEARNING",
                "ELECTRICAL AND ELECTRONICS ENGINEERING"
            ],
            variable=self.dept_var
        )
        self.dept_menu.pack(pady=10)

        self.start_btn = tk.CTkButton(self, text="Start Report Generation", command=self.start_app)
        self.start_btn.pack(pady=30)

    def start_app(self):
        """
        Validates input and signals the main loop to proceed.
        Sets `self.selected_college` and `self.selected_dept` if valid, then quits the local loop.
        """
        college = self.college_var.get()
        dept = self.dept_var.get()

        if college == "Select College" or dept == "Select Department" or college == "more coming soon" or dept == "more coming soon":
            CTkMessagebox(title="Invalid Selection", message="Please select a valid College and Department.", icon="cancel")
            return

        self.selected_college = college
        self.selected_dept = dept
        
        self.selected_college = college
        self.selected_dept = dept
        
        # Withdraw (hide) window first to stop visual updates/animations
        self.withdraw()
        
        # Quit the mainloop after a brief pause to allow pending tasks to clear
        self.after(100, self.quit)

# =================================================================================================
#                                     MAIN EXECUTION
# =================================================================================================

def main():
    """
    Main entry point.
    1. Runs StartScreen.
    2. Checks for valid selection.
    3. Launches main GUI.
    """
    tk.set_appearance_mode("dark")
    tk.set_default_color_theme("dark-blue")
    
    app = StartScreen()
    app.mainloop()
    
    # Check if we should launch the main GUI
    if hasattr(app, "selected_college") and app.selected_college:
        college = app.selected_college
        dept = app.selected_dept
        app.destroy() # Clean up start screen resources
        
        # Import gui using absolute path now that sys.path is set
        import app.frontend.gui as gui
        gui.launch_gui(college, dept)

# =================================================================================================

if __name__ == "__main__":
    main()
