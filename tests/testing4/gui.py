import customtkinter as ctk
from hf import WordDocGenerator
from pathlib import Path

# Create document generator instance
DOC_PATH = Path.cwd() / "tests" / "testing4" / "testHF.docx"
generator = WordDocGenerator(DOC_PATH)

# GUI setup
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

app = ctk.CTk()
app.title("Project Info Inserter")
app.geometry("500x300")

# ------------------ EVENT HANDLER ------------------ #
def on_done():
    title = title_entry.get().strip()
    year = year_entry.get().strip()

    if not title or not year:
        print("⚠️ Both title and year are required.")
        return

    # Create a sample data_dict like the full app
    data_dict = {
        "Project Title": title,
        "Year": year
    }

    # Pass to updated bookmark replacer
    generator.replace_bookmarks(data_dict)
    generator.save()
    print("✅ Document updated successfully.")

# ------------------ UI ELEMENTS ------------------ #
title_label = ctk.CTkLabel(app, text="Enter Project Title:")
title_label.pack(pady=(15, 5))

title_entry = ctk.CTkEntry(app, width=300)
title_entry.pack()

year_label = ctk.CTkLabel(app, text="Enter Year:")
year_label.pack(pady=(15, 5))

year_entry = ctk.CTkEntry(app, width=300)
year_entry.pack()

button = ctk.CTkButton(app, text="Done", command=on_done)
button.pack(pady=25)

# ------------------ START APP ------------------ #
app.mainloop()
