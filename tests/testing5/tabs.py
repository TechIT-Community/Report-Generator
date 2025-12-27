import customtkinter as ctk

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")


class WizardApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Wizard")
        self.geometry("900x550")

        self.container = ctk.CTkFrame(self)
        self.container.pack(fill="both", expand=True, padx=20, pady=20)

        self.pages = [
            PageOne(self.container),
            PageTwo(self.container),
            PageThree(self.container),
        ]

        for p in self.pages:
            p.place(relwidth=1, relheight=1)

        self.current = 0
        self.pages[0].tkraise()

        nav = ctk.CTkFrame(self)
        nav.pack(pady=10)

        ctk.CTkButton(nav, text="Previous", command=self.prev).pack(side="left", padx=5)
        ctk.CTkButton(nav, text="Next", command=self.next).pack(side="left", padx=5)

    def next(self):
        if self.current < len(self.pages) - 1:
            self.current += 1
            self.pages[self.current].tkraise()

    def prev(self):
        if self.current > 0:
            self.current -= 1
            self.pages[self.current].tkraise()


class PageOne(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)

        ctk.CTkLabel(self, text="Entry 1").pack(anchor="w")
        ctk.CTkEntry(self).pack(fill="x")

        ctk.CTkLabel(self, text="Entry 2").pack(anchor="w", pady=(10, 0))
        ctk.CTkEntry(self).pack(fill="x")


class PageTwo(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)

        header = ctk.CTkFrame(self)
        header.pack(fill="x")

        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(fill="both", expand=True, pady=10)

        ctk.CTkButton(header, text="+", width=30, command=self.add_tab).pack(side="right")

        self.tab_data = []  # list of dicts: {"name": str, "a": str, "b": str}
        self.counter = 0

        for _ in range(3):
            self.add_tab()

    def add_tab(self):
        self.counter += 1
        name = f"Tab {self.counter}"

        data = {"name": name, "a": "", "b": ""}
        self.tab_data.append(data)

        self._render_tabs(select=name)

    def rename_current_tab(self):
        old = self.tabview.get()
        idx = next(i for i, d in enumerate(self.tab_data) if d["name"] == old)

        dialog = ctk.CTkInputDialog(
            title="Rename Tab",
            text="New name:",
        )
        new = dialog.get_input()

        if not new or new == old:
            return

        self.tab_data[idx]["name"] = new
        self._render_tabs(select=new)

    def remove_current_tab(self):
        if not self.tab_data:
            return

        name = self.tabview.get()
        self.tab_data = [d for d in self.tab_data if d["name"] != name]

        select = self.tab_data[0]["name"] if self.tab_data else None
        self._render_tabs(select=select)

    def _render_tabs(self, select=None):
        self.tabview.destroy()
        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(fill="both", expand=True, pady=10)

        for d in self.tab_data:
            tab = self.tabview.add(d["name"])

            top = ctk.CTkFrame(tab, fg_color="transparent")
            top.pack(anchor="e")

            ctk.CTkButton(top, text="Rename", width=70,
                          command=self.rename_current_tab).pack(side="left", padx=5)
            ctk.CTkButton(top, text="✕", width=30,
                          command=self.remove_current_tab).pack(side="left")

            ctk.CTkLabel(tab, text="Field A").pack(anchor="w")
            a = ctk.CTkEntry(tab)
            a.insert(0, d["a"])
            a.pack(fill="x")
            a.bind("<KeyRelease>", lambda e, d=d, w=a: d.update(a=w.get()))

            ctk.CTkLabel(tab, text="Field B").pack(anchor="w", pady=(10, 0))
            b = ctk.CTkEntry(tab)
            b.insert(0, d["b"])
            b.pack(fill="x")
            b.bind("<KeyRelease>", lambda e, d=d, w=b: d.update(b=w.get()))

        if select:
            self.tabview.set(select)


class PageThree(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)

        ctk.CTkLabel(self, text="Final Entry 1").pack(anchor="w")
        ctk.CTkEntry(self).pack(fill="x")

        ctk.CTkLabel(self, text="Final Entry 2").pack(anchor="w", pady=(10, 0))
        ctk.CTkEntry(self).pack(fill="x")


if __name__ == "__main__":
    WizardApp().mainloop()
