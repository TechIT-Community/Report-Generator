import customtkinter as ctk

# --- Global Configuration ---
ctk.set_appearance_mode("dark")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("dark-blue")  # Themes: "blue" (standard), "green", "dark-blue"

class WizardApp(ctk.CTk):
    """
    The main Application Window. 
    It acts as a 'Controller' to switch between different Page frames (Wizard style).
    """
    def __init__(self):
        super().__init__()
        self.title("Wizard Application")
        self.geometry("900x600")

        # 1. Main Container: All pages will be placed inside this frame.
        self.container = ctk.CTkFrame(self)
        self.container.pack(fill="both", expand=True, padx=20, pady=20)

        # 2. Page Initialization: Create instances of each page class.
        self.pages = [
            PageOne(self.container),
            PageTwo(self.container),
            PageThree(self.container),
        ]

        # 3. Layout Strategy: Use 'place' to stack pages on top of each other.
        # Calling .tkraise() on a frame will bring it to the front of the stack.
        for p in self.pages:
            p.place(relwidth=1, relheight=1)

        self.current = 0
        self.pages[0].tkraise() # Show the first page initially

        # 4. Navigation Bar: Fixed at the bottom of the window.
        nav = ctk.CTkFrame(self)
        nav.pack(pady=10, fill="x")

        # Center buttons inside the navigation frame
        btn_container = ctk.CTkFrame(nav, fg_color="transparent")
        btn_container.pack(expand=True)

        ctk.CTkButton(btn_container, text="Previous", command=self.prev).pack(side="left", padx=5)
        ctk.CTkButton(btn_container, text="Next", command=self.next).pack(side="left", padx=5)

    def next(self):
        """Moves to the next page in the list."""
        if self.current < len(self.pages) - 1:
            self.current += 1
            self.pages[self.current].tkraise()

    def prev(self):
        """Moves to the previous page in the list."""
        if self.current > 0:
            self.current -= 1
            self.pages[self.current].tkraise()


class PageOne(ctk.CTkFrame):
    """Simple form-style page."""
    def __init__(self, master):
        super().__init__(master)
        ctk.CTkLabel(self, text="Step 1: Basic Information", font=("Arial", 20, "bold")).pack(pady=10)
        
        ctk.CTkLabel(self, text="Full Name").pack(anchor="w", padx=20)
        ctk.CTkEntry(self).pack(fill="x", padx=20, pady=(0, 10))

        ctk.CTkLabel(self, text="Email Address").pack(anchor="w", padx=20)
        ctk.CTkEntry(self).pack(fill="x", padx=20)


# ---------------- MULTI-ROW TAB PAGE (The Complex Part) ----------------
class PageTwo(ctk.CTkFrame):
    """
    A dynamic page that allows users to add/remove/rename tabs.
    Tabs are displayed in a grid (multi-row) within the top bar.
    """
    TABS_PER_ROW = 5

    def __init__(self, master):
        super().__init__(master)

        # --- Layout Components ---
        # Top Bar: Contains the grid of Tab Buttons and the [+] button
        top = ctk.CTkFrame(self)
        top.pack(fill="x", padx=5, pady=5)

        # Tab Bar: A sub-frame where buttons are gridded
        self.tab_bar = ctk.CTkFrame(top, fg_color="transparent")
        self.tab_bar.pack(side="left", fill="x", expand=True)

        # Plus Button: Fixed to the right side
        self.add_btn = ctk.CTkButton(
            top, text="+", width=40, font=("Arial", 16, "bold"), 
            command=self.add_tab
        )
        self.add_btn.pack(side="right", padx=10)

        # Content Area: Where the active tab's frame is packed/unpacked
        self.content_container = ctk.CTkFrame(self)
        self.content_container.pack(fill="both", expand=True, pady=10, padx=10)

        # --- Data Storage ---
        self.tabs = []            # Stores dicts: {"name": str, "frame": CTkFrame, "data": dict}
        self.active_tab = None    # Pointer to the currently selected tab dict
        self.counter = 0          # Used to generate unique default names (Tab 1, Tab 2...)

        # Create a few tabs by default
        for _ in range(3):
            self.add_tab()

    # ---------- TAB OPERATIONS ----------

    def add_tab(self):
        """Creates a new tab data structure and its associated UI frame."""
        self.counter += 1
        tab_id = f"Chapter {self.counter}"
        
        # Define the tab object (Dictionary)
        tab = {
            "name": tab_id,
            "frame": None, # Will be created below
            "data": {"field_a": "", "field_b": ""}
        }

        # Create the UI Frame that belongs to this specific tab
        tab["frame"] = self.create_tab_ui(tab)
        self.tabs.append(tab)

        # Switch focus to the newly created tab
        self.set_active(tab)

    def remove_tab(self, tab):
        """Removes a tab from the list and destroys its frame."""
        if tab is self.active_tab:
            tab["frame"].pack_forget()

        tab["frame"].destroy()
        self.tabs.remove(tab)

        # Logic to decide which tab to show next
        if self.tabs:
            self.set_active(self.tabs[0])
        else:
            self.active_tab = None
            self.render_tab_bar()

    def rename_tab(self, tab):
        """Opens a dialog to rename the tab and refreshes the UI."""
        dialog = ctk.CTkInputDialog(title="Rename", text=f"Enter name for {tab['name']}:")
        new_name = dialog.get_input()

        if new_name:
            tab["name"] = new_name
            # Re-render to show the new name on the button while keeping the highlight
            self.render_tab_bar()

    def set_active(self, tab):
        """The 'Engine' of the tab system. Handles visibility and highlighting."""
        # 1. Hide the currently visible tab frame
        if self.active_tab:
            self.active_tab["frame"].pack_forget()

        # 2. Update the reference to the new active tab
        self.active_tab = tab
        
        # 3. Show the new frame
        if tab:
            tab["frame"].pack(fill="both", expand=True)

        # 4. CRITICAL: Redraw all tab buttons to update their colors
        self.render_tab_bar()

    # ---------- RENDERING ----------

    def render_tab_bar(self):
        """Clears the tab_bar frame and redraws buttons based on the current 'tabs' list."""
        # Clean up old widgets
        for widget in self.tab_bar.winfo_children():
            widget.destroy()

        # Iterate through all tab objects and create a button for each
        for index, tab in enumerate(self.tabs):
            # Calculate grid position based on TABS_PER_ROW
            row = index // self.TABS_PER_ROW
            col = index % self.TABS_PER_ROW

            # Determine colors based on active state
            is_active = (tab is self.active_tab)
            # Use theme blue for active, dark gray for inactive
            bg_color = "#1f538d" if is_active else "#333333"
            hover_color = "#2b71ba" if is_active else "#444444"

            btn = ctk.CTkButton(
                self.tab_bar,
                text=tab["name"],
                width=140,
                fg_color=bg_color,
                hover_color=hover_color,
                # Pass the current 'tab' to the command via default argument t=tab
                command=lambda t=tab: self.set_active(t)
            )
            btn.grid(row=row, column=col, padx=5, pady=5)

    def create_tab_ui(self, tab):
        """Generates the internal UI widgets for a specific tab frame."""
        frame = ctk.CTkFrame(self.content_container)

        # Tab Toolbar (Rename and Delete buttons)
        toolbar = ctk.CTkFrame(frame, fg_color="transparent")
        toolbar.pack(fill="x", anchor="e")

        ctk.CTkButton(
            toolbar, text="Rename", width=60, height=24, font=("Arial", 11),
            command=lambda t=tab: self.rename_tab(t)
        ).pack(side="right", padx=5)

        ctk.CTkButton(
            toolbar, text="✕", width=24, height=24, fg_color="#a32a2a", hover_color="#822121",
            command=lambda t=tab: self.remove_tab(t)
        ).pack(side="right")

        # Input Fields
        ctk.CTkLabel(frame, text=f"Data for {tab['name']}", font=("Arial", 16, "italic")).pack(pady=10)

        ctk.CTkLabel(frame, text="Variable A").pack(anchor="w", padx=20)
        ent_a = ctk.CTkEntry(frame)
        ent_a.pack(fill="x", padx=20, pady=(0, 10))
        # Bind typing event to save data into the tab dictionary
        ent_a.bind("<KeyRelease>", lambda e, t=tab, w=ent_a: t["data"].update({"field_a": w.get()}))

        ctk.CTkLabel(frame, text="Variable B").pack(anchor="w", padx=20)
        ent_b = ctk.CTkEntry(frame)
        ent_b.pack(fill="x", padx=20)
        ent_b.bind("<KeyRelease>", lambda e, t=tab, w=ent_b: t["data"].update({"field_b": w.get()}))

        return frame


class PageThree(ctk.CTkFrame):
    """Final summary page."""
    def __init__(self, master):
        super().__init__(master)
        ctk.CTkLabel(self, text="Step 3: Review & Submit", font=("Arial", 20, "bold")).pack(pady=20)
        ctk.CTkCheckBox(self, text="I agree to the terms and conditions").pack(pady=10)
        ctk.CTkButton(self, text="Complete Setup", fg_color="green", hover_color="darkgreen").pack(pady=20)


if __name__ == "__main__":
    # Initialize and run the application
    app = WizardApp()
    app.mainloop()