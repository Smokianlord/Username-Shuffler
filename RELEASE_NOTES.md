# Username Shuffler v3.0.0 Release Notes

We are thrilled to announce **Username Shuffler v3.0.0**, a complete overhaul of the desktop application! This major release brings a modern UI with a top action toolbar, instantaneous launch speeds, memory-efficient streaming file parsers, advanced shuffle modes, comprehensive output formatting, and live Dark/Light theme switching.

---

### 🌟 Key Highlights & What's New

#### 🖥️ 1. Modern Desktop UI with Top Action Toolbar
- **Modern Action Toolbar Ribbon**: Quick access to essential functions:
  - `📂 Open File`: Open any dataset from any directory.
  - `🔄 Reload`: Hot reload current file from disk.
  - `🔀 Shuffle`: Instant manual re-draw.
  - `📋 Copy Results`: One-click copy with visual toast feedback.
  - `💾 Export...`: Save shuffled output to `.txt`, `.csv`, or `.xlsx`.
  - `🌓 Theme`: Live switch between Light and Dark modes.
  - `📄 Dataset Status Pill`: Displays current active file and loaded user count.
- **Native Menu Bar**: Standard `File`, `Edit`, `Shuffle`, `View`, and `Help` menus with full keyboard shortcuts.
- **Modern Elevated Cards**: Refined borders, crisp typography, clean inputs, and adaptive layouts.
- **High-DPI Awareness**: Crystal-clear text and sharp icons on 1080p, 2K, and 4K displays.

#### ⚡ 2. Launchup & Performance Optimizations
- **Instant Launch**: Removed blocking calls and repetitive Win32 icon hooks. Window opens immediately.
- **Streaming XML Parser (`xml.etree.ElementTree.iterparse`)**: Parses large `.xlsx` / `.xlsm` workbooks with minimal RAM and near-zero CPU latency.
- **Thread-Safe Asynchronous Loading (`queue.Queue`)**: Background file loading keeps the interface 100% responsive without freezing or thread collision.
- **Compact Standalone EXE**: Optimized build produces a portable standalone executable of only ~11 MB with zero external runtime dependencies.

#### 🔀 3. Advanced Shuffle Modes
- **Unique Picks (Standard)**: Random selection without replacement (no duplicate winners).
- **With Replacement**: Allows repeats and enables picking counts larger than the total pool.
- **Elimination Draw (Giveaway / Raffle / Tournament Mode)**: Eliminates drawn usernames across consecutive rounds with live pool countdown (`X remaining in pool`) and a 1-click `↺ Reset Pool` button.
- **Quick Pick Chips**: Instant chips for `[1]`, `[5]`, `[10]`, `[25]`, `[50]`, and `[All]`.

#### 🛠️ 4. Rich Output Customization
- **Flexible Delimiters**: Choose between **Newline** (one per line), **Space**, **Comma** (`, `), **Semicolon** (`; `), or a **Custom delimiter**.
- **Social Handles & Affixes**: Add Prefix (e.g. `@` for Twitter/Instagram/Discord) with smart duplicate prevention, and Suffixes.
- **Numbered Lists**: Toggle numbered formatting (`1. @alice`, `2. @bob`).
- **Case Conversion**: Keep original casing, or transform to `lowercase`, `UPPERCASE`, or `Title Case`.

#### 📂 5. Multi-Format Support & Smart Header/Column Detection
- Supports **Excel (`.xlsx`, `.xlsm`)**, **CSV (`.csv`)**, and **Plain Text (`.txt`)** files.
- **Intelligent Header Detection**: Automatically identifies header rows (`Username`, `Handle`, `Name`, etc.) and excludes them so headers are never shuffled as participants.
- **Multi-Column Column Selector**: If a spreadsheet has multiple columns (e.g. ID, Username, Email), effortlessly switch between columns via the active column dropdown.
- **Open Any File**: Browse and open files from anywhere on your PC, with automatic remembering of recent files.

#### ➕ 6. Dataset Manager & Tools
- **Add Usernames Tab**: Bulk paste box supporting line breaks, commas, and semicolons, with automatic deduplication against existing pool and 1-click append to file.
- **Search Dataset Tab**: Real-time search to instantly verify if a username exists in the dataset.
- **Draw History Tab**: Timestamped log of previous draws during your session with 1-click recall.

#### 🌓 7. Live Dark & Light Theme System
- Fully integrated theme engine with custom color palettes for both Dark Mode and Light Mode.
- Automatically detects Windows system theme on first launch, or toggle anytime with `Ctrl + D` or the top toolbar.

---

### 🐛 Resolved Bugs & Issues
1. **Header Row Bug**: Fixed header row words like "Username" or "Account" being shuffled into results.
2. **Multi-Column Bleed**: Fixed multi-column spreadsheets merging emails/IDs into usernames.
3. **Missing File Dialog**: Resolved inability to open files outside the app folder.
4. **Delimiters**: Resolved usernames only being space-separated; full support for newlines, commas, and numbering.
5. **Deleted Script References**: Removed broken references to deleted `.bat` files in documentation and provided clean `build.py` / `build.bat`.
6. **Thread Crash on Windows**: Replaced insecure thread calls with thread-safe `queue.Queue` dispatching.
7. **DPI Blurriness**: Enabled native Windows High-DPI awareness for crisp rendering.
8. **Resource Leaks**: Cleaned up icon handle allocation on Windows.

---

### ⌨️ Keyboard Shortcuts
| Shortcut | Action |
| :--- | :--- |
| **`Ctrl + O`** | Open Dataset File |
| **`Ctrl + R`** or **`F5`** | Reload Active File |
| **`Ctrl + Enter`** or **`Enter`** | Shuffle Now |
| **`Ctrl + C`** | Copy Formatted Results to Clipboard |
| **`Ctrl + E`** or **`Ctrl + S`** | Export Results (TXT, CSV, XLSX) |
| **`Ctrl + D`** | Toggle Dark / Light Theme |
| **`Ctrl + K`** | Clear Results |

---

### 📦 Release Assets
- `Username-Shuffler-v3.0.0-Windows.zip`: Standalone Windows bundle (executable, sample dataset, quick start guide).
- `Username-Shuffler.exe`: Standalone portable executable.
- `Username Dataset.xlsx`: Sample dataset.
