# Excel Column Mapper & Merger (Streamlit)

Merge multiple Excel files into a single CSV with standardized column names — even when headers differ slightly. Includes fuzzy header matching, an editable mapping table, a `source_file` column, detailed logs, and local folder pickers.

## Features
- Select **input folder** (recursively optional) and **output folder**
- Auto-suggest mappings using fuzzy matching (configurable threshold)
- Edit mappings in a table; group many headers → one standardized name
- Merge to CSV with a `source_file` column
- Logs and validation per file + error reporting
- 100% local processing


---

## Run on macOS
1. Install **Python 3.10+** (`python3 --version`).
2. In Terminal:
   ```bash
   cd /path/to/excel-merger
   chmod +x run_mac.command
   ./run_mac.command

   # Excel Column Mapper & Merger (Streamlit)



Merge multiple Excel files into a single CSV with standardized column names — even when headers differ slightly. Includes fuzzy header matching, an editable mapping table, a `source_file` column, detailed logs, and local folder pickers.

## Features
- Select **input folder** (recursively optional) and **output folder**
- Auto-suggest mappings using fuzzy matching (configurable threshold)
- Edit mappings in a table; group many headers → one standardized name
- Merge to CSV with a `source_file` column
- Logs and validation per file + error reporting
- 100% local processing

---

## Run on macOS
1. Install **Python 3.10+** (`python3 --version`).
2. In Terminal:
   ```bash
   cd /path/to/excel-merger
   chmod +x run_mac.command

## Run on Windows
1.	Install Python 3.10+ and ensure py is on PATH.
2.	Double-click run_windows.bat.
- If a Windows Defender prompt appears, choose More info → Run anyway (if applicable in your environment).

---

## Deployment Guide (Windows)

This guide will walk you through setting up and running the Excel Merger application on a Windows machine using either the command line or Visual Studio Code.

### Prerequisites

Before you begin, ensure you have the following installed:

1.  **Python:** Download and install a recent version of Python from [python.org](https://www.python.org/downloads/).
    *   **Important:** During installation, make sure to check the box that says **"Add Python to PATH"**.
2.  **Git:** (Optional, for cloning) Download and install Git from [git-scm.com](https://git-scm.com/download/win) to easily download the project.

---

### Setup & Installation

Follow these steps to get the application ready. You only need to do this once.

1.  **Download the Project:**
    *   **With Git:** Open a Command Prompt and run:
        ```bash
        git clone https://github.com/your-username/Excel-Merger.git
        cd Excel-Merger
        ```
        *(Replace `your-username/Excel-Merger.git` with the actual repository URL)*
    *   **Without Git:** Download the project as a ZIP file from GitHub and extract it to a folder on your computer.

2.  **Run the Setup Script:**
    The project includes a file called `run_windows.bat` that automates the entire setup process. It will:
    *   Create an isolated Python "virtual environment" in a `.venv` folder.
    *   Activate this environment.
    *   Install all the required Python packages (like Streamlit and Pandas).

---

### How to Run the Application

#### Option 1: From File Explorer (Easiest Method)

1.  Navigate to the project folder where you extracted the files.
2.  Double-click the **`run_windows.bat`** file.

A command prompt window will appear, set up the environment, and then launch the application in your default web browser.

#### Option 2: From the VS Code Terminal

1.  Open the project folder in Visual Studio Code.
2.  Open a new terminal by pressing **`Ctrl + `** (the backtick key, next to the `1` key) or by going to `Terminal > New Terminal`.
3.  In the terminal window that appears, type the following command and press Enter:

    ```powershell
    .\run_windows.bat
    ```

The application will start and open in your web browser.

---

### How to Use the App

1.  **Place Files:** Put the Excel or CSV files you want to merge into the `input` folder.
2.  **Configure Settings:** Use the **Utilities** section to point the app to your input folder if it's not the default one.
3.  **Create Groups:**
    *   Go to the **Mapping Builder**.
    *   Give your output column a name (e.g., "Standardized Email").
    *   For each file listed in the table, use the dropdown to select the column that corresponds to your group (e.g., select "Email Address" from one file, "Contact" from another).
    *   Click **Save group**. Repeat for all the columns you want to merge.
4.  **Merge:** Once you have created all your groups, click the **Merge & Export** button at the bottom. Your merged file will be saved in the `output` folder.

### Stopping the Application

To stop the app, simply close the black command prompt window. If you are running it in the VS Code terminal, click inside the terminal and press **`Ctrl + C