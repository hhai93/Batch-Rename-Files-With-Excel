# Batch Rename Files With Excel

![GitHub](https://img.shields.io/badge/license-MIT-blue.svg) ![GitHub last commit](https://img.shields.io/github/last-commit/hhai93/Batch-Rename-Files-With-Excel)

A VBA script to batch rename files in a folder with a user-friendly interface, folder picker, preview functionality, and file type filtering.

---

## ✨ Features
- 📁 Select folder via a dialog box.
- ✏️ Customize renaming with prefix, suffix, and optional numbering.
- 🔍 Filter files by extension (e.g., `.pdf` only).
- 👀 Preview new file names in Excel before renaming.
- 📋 Logs old and new names in the active sheet.
- ⚠️ Error handling for failed renames.

## 🛠️ Prerequisites
- 💻 Microsoft Excel (2010 or later) with VBA enabled.
- 📂 A folder containing files to rename.

---

## 🚀 How to Use

### 1. Prepare Your Folder
- Create a folder with files (e.g., `C:\TestFolder\`).

### 2. Add the VBA Script
- Open Excel and press `Alt + F11` to access the VBA Editor.
- **Add Module**:
  - Go to **Insert** > **Module**, paste the code from [`BatchRename.vba`](BatchRename.vba).
- **Add UserForm**:
  - Go to **Insert** > **UserForm**, name it `BatchRenameForm`.
  - Add controls as described in [`BatchRenameForm.vb`](BatchRenameForm.vb) under "UserForm Layout".
  - Paste the code from "UserForm Code" into the UserForm's code window.
- Save the file as `.xlsm` (macro-enabled).

### 3. Run the Script
- Press `Alt + F8`, select `BatchRename`, and click **Run**.
- In the UserForm:
  - Click **Browse** to select a folder.
  - Enter a prefix (e.g., `New_`), suffix (e.g., `_v1`), and file type (e.g., `.pdf`) if needed.
  - Check **Add Numbering** for sequential numbers.
  - Click **Preview** to see new names in the sheet.
  - Click **Rename** to apply changes.
- Check the folder for renamed files! 🎉

---

## 💡 Code Explanation
- **`BatchRename.vba`**:
  - Launches the UserForm and includes a folder picker helper function.
- **`BatchRenameForm.vb`**:
  - **UserForm**: Interface for folder selection and renaming options.
  - **Preview**: Generates new names and logs them in the sheet.
  - **Rename**: Applies changes based on the preview.
- **Key Features**:
  - Uses `Dir` to list files dynamically.
  - Supports prefix, suffix, and numbering customization.
  - Filters files by extension if specified.

## ⚙️ Customization
- Modify default values in `UserForm_Initialize` (e.g., change `txtPrefix.Value`).
- Adjust the `newFileName` logic in `btnPreview_Click` for different rules.

## ⚠️ Notes
- 💾 Back up files before renaming to avoid accidental overwrites.
- 📍 Ensure the folder path is valid when previewing.
- 🚨 Errors (e.g., file in use) are logged in the "New Name" column.
