# Excel Column Mapper & Merger

A powerful Streamlit web application designed to intelligently merge data from multiple Excel (`.xlsx`, `.xls`) and CSV files. This tool excels at handling files with inconsistent column names by allowing you to create "groups" that map different header variants to a single, standardized output column.

Save your complex mapping configurations as reusable **presets** to automate recurring merge tasks with a single click.

## Table of Contents

- [Key Features](#key-features)
- [Quick Start Guide](#quick-start-guide)
  - [On Windows](#on-windows)
  - [On macOS / Linux](#on-macos--linux)
- [Detailed User Guide](#detailed-user-guide)
  - [Setup](#setup)
  - [The Main Interface](#the-main-interface)
  - [Core Workflow: Merging Files](#core-workflow-merging-files)
  - [Using Presets (For Recurring Tasks)](#using-presets-for-recurring-tasks)
- [Deployment with VS Code](#deployment-with-vs-code)

---

## Key Features

*   **Intelligent Grouping:** Map various source column names (e.g., "Email", "Email Address", "Contact") to one output column (e.g., "Standardized Email").
*   **Preset Management:** Save, load, duplicate, and set default mapping configurations as named presets to streamline repeated tasks.
*   **Advanced Mapping Builder:**
    *   Fuzzy-logic "Auto-Match" to suggest column mappings.
    *   Interactive table to manually map columns.
    *   Filters to quickly find files that still need mapping or to search by name.
*   **Flexible File Handling:**
    *   Process files from a specified folder, with optional recursive search.
    *   Optionally treat individual sheets within an Excel file as separate data units.
*   **Validation & Preview:**
    *   Review header coverage across all files before merging.
    *   Validate your mappings to see which files are missing a column for a specific group.
*   **Traceability:** Optionally add source file/sheet columns to the output for row-by-row traceability.
*   **Cross-Platform:** Includes simple run scripts for both Windows and macOS/Linux.

---

## Quick Start Guide

### Prerequisites

*   **Python 3.10+** installed. During installation, ensure you check **"Add Python to PATH"** (on Windows).

### On Windows

1.  Download and unzip the project.
2.  Double-click the **`run_windows.bat`** file.
3.  The app will open in your web browser.

### On macOS / Linux

1.  Download and unzip the project.
2.  Open your **Terminal**.
3.  Navigate into the project folder (e.g., `cd Downloads/Excel-Merger`).
4.  Run the command: `sh run_mac_linux.sh`
5.  The app will open in your web browser.

---

## Detailed User Guide

### 1. Setup

The included `run` scripts (`run_windows.bat` and `run_mac_linux.sh`) automate the entire setup process. The first time you run the script, it will:
1.  Create an isolated Python virtual environment in a `.venv` folder.
2.  Activate the environment.
3.  Install all required packages (Streamlit, Pandas, etc.).

Subsequent runs will be much faster as they will use the existing environment.

### 2. The Main Interface

The application is organized into collapsible sections:

*   **Utilities:** Configure input/output folders and core settings.
*   **Group Presets:** Save and load your entire mapping configuration.
*   **File and Sheet Selection:** View all detected files/sheets and choose which ones to include in the merge.
*   **Header Preview:** See a summary of all unique column headers found in the included files.
*   **Mapping Builder:** The main workspace for creating your column groups.
*   **Saved Header Groups:** View, edit, or remove the groups you've created.
*   **Validate:** Check your work before the final merge.
*   **Merge & Export:** The final button to perform the merge and download the result.

### 3. Core Workflow: Merging Files

#### Step 1: Configure Settings & Select Files

1.  Use the **Utilities** section to point the app to your input folder if it's not the default `input/` directory.
2.  In the **File and Sheet Selection** section, ensure all the files you want to merge are checked in the "Included units" list.

#### Step 2: Create Column Groups

A "group" defines how one or more source columns are mapped to a single output column.

1.  Expand the **Mapping Builder**.
2.  **Name the Output Column:** In the "Output column name" field, enter the name you want for the final column (e.g., "Contact Email").
3.  **Map Columns:** For each file listed in the table, use the dropdown to select the corresponding source column.
    *   **Tip (Auto-Match):** To speed things up, type a keyword (like "email") into the "Auto-match pattern" field and click "Apply Auto-Match". The app will attempt to find and select the best match in each file.
    *   **Tip (Filtering):** Use the "Show only files without a selection" toggle to hide files you've already mapped. Use the "Filter files by name" box to find specific files quickly.
4.  **Save the Group:** Once you are happy with the mapping for this group, click **"💾 Save group"**.
5.  Repeat this process for every standardized column you want in your final output file.

#### Step 3: Validate and Merge

1.  Expand the **Validate** section and click **"🧪 Validate"**. This shows a table of every group and file, highlighting any "MISSING" mappings.
2.  If there are missing mappings, you can either go back and fix them or check the "Understood — proceed with blanks" box.
3.  Click the **"✅ Merge & Export"** button at the bottom of the page. Your merged file will be generated and a download button will appear.

### 4. Using Presets (For Recurring Tasks)

Presets save your entire collection of saved groups, allowing you to instantly load a complex configuration.

*   **To Save a Preset:**
    1.  Create all your desired groups in the Mapping Builder.
    2.  Expand the **Group Presets** section.
    3.  Under "Save Current Groups", type a descriptive name for your preset.
    4.  Click **"💾 Save Current Groups as Preset"**.

*   **To Load a Preset:**
    1.  Expand the **Group Presets** section.
    2.  Select a preset from the dropdown menu.
    3.  Click **"📂 Load Selected Preset"**. All your saved groups will be instantly loaded.

*   **To Edit a Preset:**
    1.  Load the preset.
    2.  Make your changes in the Mapping Builder or Saved Groups sections.
    3.  Save it again with the **exact same name** to overwrite it.

*   **To Duplicate a Preset:**
    1.  Select the preset you want to copy.
    2.  Click **"📠 Duplicate Preset"**. A copy will be created with a `_copy` suffix.

*   **To Set a Default:**
    1.  Select a preset.
    2.  Click **"⭐ Set as Default"**. This preset will now be loaded automatically every time you start the app.

---

## Deployment with VS Code

1.  Open the project folder in Visual Studio Code.
2.  Open a new terminal (`Terminal > New Terminal` or `Ctrl+`\`).
3.  In the terminal, run the appropriate script for your OS:
    *   **Windows:** `.\run_windows.bat`
    *   **macOS/Linux:** `sh run_mac_linux.sh`
4.  The application will start, and VS Code may prompt you to open the URL in your web browser.