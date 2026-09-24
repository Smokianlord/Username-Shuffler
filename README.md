# Username Shuffler v3.0.0

[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Platform Windows](https://img.shields.io/badge/platform-windows-0078D6.svg)](https://github.com/Smokianlord/Username-Shuffler/releases)
[![Release](https://img.shields.io/badge/release-v3.0.0-green.svg)](https://github.com/Smokianlord/Username-Shuffler/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A modern, high-performance desktop application for instantly shuffling, formatting, managing, and exporting usernames from Excel (`.xlsx`, `.xlsm`), CSV, and Plain Text (`.txt`) files.

<p align="center">
  <img width="800" alt="Username Shuffler" src="https://github.com/user-attachments/assets/ddc83302-2534-4fd4-a3e9-014d77f9f6ce" />
</p>

---

## ✨ Features

- **🖥️ Modern Action Toolbar Ribbon**: Quick access to Open File, Reload, Shuffle, Copy, Export, and Dark/Light mode toggle.
- **⚡ Instant Launch & High Performance**: Zero blocking startup, streaming XML parser for large spreadsheets, and multi-threaded asynchronous loading.
- **🔀 Advanced Shuffle Modes**:
  - **Unique (Standard)**: Select random items without replacement.
  - **With Replacement**: Allow repeat picks or pick more items than the total pool.
  - **Elimination Draw (Giveaway / Raffle)**: Shrinks the pool as you draw across rounds until pool reset.
  - **Quick Amount Chips**: Pick `1`, `5`, `10`, `25`, `50`, or `All` in one click.
- **🛠️ Rich Output Formatting**:
  - Separators: Newline (one per line), Space, Comma (`, `), Semicolon (`; `), or Custom.
  - Social Handle Prefix (e.g. `@`) and Suffix support.
  - Numbered list formatting (`1. @alice`, `2. @bob`).
  - Text transformation: Original, lowercase, UPPERCASE, Title Case.
- **📂 Multi-Format & Column Detection**:
  - Works with `.xlsx`, `.xlsm`, `.csv`, and `.txt` files.
  - Intelligent header row filtering (skips "Username", "Name", "Handle" headers automatically).
  - Multi-column selector dropdown for multi-column spreadsheets.
- **📋 Dataset Management Tools**:
  - **Add Usernames**: Bulk paste new usernames with real-time duplicate detection and save directly to file.
  - **Search Dataset**: Instant search to check if a user is in your dataset.
  - **Draw History**: Timestamped history of all draws during your session with 1-click recall.
- **🌓 Dark & Light Theme**: Seamless live switching with Windows system theme detection.
- **⌨️ Keyboard Shortcuts**: Full keyboard control for power users.

---

## 🚀 How to Use

### Option 1: Standalone Executable (Recommended)

1. Download `Username-Shuffler-v3.0.0-Windows.zip` from the [Latest Release](https://github.com/Smokianlord/Username-Shuffler/releases).
2. Extract the ZIP file anywhere on your computer.
3. Double-click `Username-Shuffler.exe` to run.
4. The app will automatically detect any spreadsheet in the folder (like `Username Dataset.xlsx`), or click **📂 Open File** to browse any file on your computer.
5. Type the number of usernames you want (or click a quick chip).
6. Click **📋 Copy Results** or press `Ctrl + C`!

### Option 2: Run from Python Source

Requirements: Python 3.10 or newer (no external packages required to run!).

```bash
# Clone the repository
git clone https://github.com/Smokianlord/Username-Shuffler.git
cd Username-Shuffler

# Run with standard Python (console)
python app.py

# Or run windowed without terminal (Windows)
pythonw Username-Shuffler.pyw
```

---

## ⌨️ Keyboard Shortcuts

| Shortcut | Description |
| :--- | :--- |
| **`Ctrl + O`** | Open Dataset File (`.xlsx`, `.csv`, `.txt`) |
| **`Ctrl + R`** or **`F5`** | Reload Active Dataset |
| **`Ctrl + Enter`** or **`Enter`** | Shuffle / Draw Usernames |
| **`Ctrl + C`** | Copy Formatted Result to Clipboard |
| **`Ctrl + E`** or **`Ctrl + S`** | Export Results to File |
| **`Ctrl + D`** | Toggle Dark / Light Theme |
| **`Ctrl + K`** | Clear Results Box |

---

## 📊 Dataset File Formats

Username Shuffler supports multiple formats out of the box:

### 1. Excel (`.xlsx`, `.xlsm`)
Put usernames in a column. Header rows (such as `Username` or `Handle`) are automatically identified and excluded from the draw. If multiple columns exist, use the **Active Column** dropdown to select the column you wish to shuffle.

### 2. CSV (`.csv`)
Comma, tab, or semicolon-delimited values. Compatible with UTF-8, UTF-8 with BOM, and Windows CP1252.

### 3. Plain Text (`.txt`)
One username per line:
```text
@alice
@bob
@charlie
@david
```

---

## 🛠️ Building Standalone Executable

To compile your own standalone Windows `.exe` using PyInstaller:

```bash
# Install PyInstaller
pip install pyinstaller

# Run the automated build script
python build.py
```

The compiled standalone executable and release ZIP archive will be created in the `dist/` folder.

---

## 📄 License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.
