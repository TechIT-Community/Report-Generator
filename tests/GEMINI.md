# Report Generator

**Automated Report Generator** is a desktop application built with Python that simplifies the creation of formatted technical reports (specifically for BNMIT). It leverages `CustomTkinter` for a modern GUI and `pywin32` for direct interaction with Microsoft Word.

## 📂 Project Structure

*   **`app/Main.py`**: The entry point of the application. Initializes the start screen and handles the transition to the main GUI.
*   **`app/gui.py`**: Contains the main application logic and UI definitions using `CustomTkinter`. Handles user inputs, page navigation, and communication with the backend.
*   **`app/Document_Generator.py`**: The backend module responsible for automating Microsoft Word. It handles document creation, formatting, bookmark replacement, and image insertion.
*   **`app/assets/`**: Stores static assets like logos and icons.
*   **`app/reports/`**: Destination for the generated report (`template.docx`).
*   **`requirements.txt`**: List of Python dependencies.

## 🚀 Getting Started

### Prerequisites

*   **OS**: Windows (Required for `pywin32` and COM interaction).
*   **Software**: Microsoft Word must be installed.
*   **Python**: Version 3.8 or higher.

### Installation

1.  **Clone/Download** the repository.
2.  **Create a virtual environment** (recommended):
    ```powershell
    python -m venv venv
    .\venv\Scripts\activate
    ```
3.  **Install dependencies**:
    ```powershell
    pip install -r requirements.txt
    ```

### Running the Application

Execute the main script from the project root:

```powershell
python app/Main.py
```

*Note: ensure you are in the root directory or adjust the path accordingly.*

## 🛠️ Development Notes

*   **Word Automation**: The core logic relies on the Word COM interface. Ensure Word is not stuck in a dialog or error state if the app fails to launch.
*   **Bookmarks**: The document generation relies on named "bookmarks" in the Word document (`ProjectTitle`, `Chapter1Content`, etc.). `Document_Generator.py` inserts these dynamically.
*   **Images**: Images are handled by verifying filenames against a pattern (`Fig X.Y`) and inserting them into the corresponding chapter.

## 📝 Key Libraries

*   **`customtkinter`**: For the GUI.
*   **`pywin32` (`win32com.client`)**: For controlling Microsoft Word.
*   **`CTkMessagebox`**: For popup dialogs.
*   **`Pillow` (`PIL`)**: For image processing.
