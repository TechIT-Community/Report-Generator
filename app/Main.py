from tkinter import *  # Standard Tkinter for basic GUI
import customtkinter as tk  # Modern UI
from CTkMessagebox import CTkMessagebox
from pathlib import Path  # Path handling
from PIL import Image  # 🧩 Fix: Needed for CTkImage
 
# Do not import gui here to avoid Word opening prematurely

# =================================================================================

BASE_DIR = Path(__file__).resolve().parent  # Base directory of the application
ASSET_DIR = BASE_DIR / "assets"  # Directory for assets 

# =================================================================================

class StartScreen(tk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Start Report Generator")
        self.geometry("700x450")
        self.resizable(False, False)
        self.update_idletasks()
        w = 700
        h = 450
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()

        x = int((screen_w - w) / 2 + 0.05 * screen_w)  
        y = int((screen_h - h) / 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

        logo_path = ASSET_DIR / "icon.png"
        logo_image = Image.open(logo_path)
        self.logo = tk.CTkImage(light_image=logo_image, dark_image=logo_image, size=(120, 120))
        self.logo_label = tk.CTkLabel(self, image=self.logo, text="")
        self.logo_label.pack(pady=(30, 10))
        
        icon_path = str(ASSET_DIR / "icon.ico")
        self.iconbitmap(icon_path)

        self.title_label = tk.CTkLabel(self, text="REPORT GENERATOR", font=("Arial", 24, "bold"))
        self.title_label.pack(pady=(0, 20))

        self.college_var = tk.StringVar(value="Select College")
        self.college_menu = tk.CTkOptionMenu(
            self, values=["BNMIT", "more coming soon"], variable=self.college_var
        )
        self.college_menu.pack(pady=10)

        self.dept_var = tk.StringVar(value="Select Department")
        self.dept_menu = tk.CTkOptionMenu(
            self, values=[
                "COMPUTER SCIENCE AND ENGINEERING",
                "ELECTRICAL AND COMMUNICATION ENGINEERING",
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
        college = self.college_var.get()
        dept = self.dept_var.get()

        if college == "more coming soon" or dept == "more coming soon":
            CTkMessagebox(title="Not Available", message="Selected option is not yet supported.", icon="warning")
            return

        self.after(100, lambda: self._launch_and_close(college, dept))

    def _launch_and_close(self, college, dept):
        self.quit()
        self.destroy()
        import gui
        gui.launch_gui(college, dept)

# =================================================================================

def main():
    tk.set_appearance_mode("dark")
    tk.set_default_color_theme("dark-blue")
    app = StartScreen()
    app.mainloop()

# =================================================================================

if __name__ == "__main__":
    main()