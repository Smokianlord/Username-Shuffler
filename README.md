# Username Shuffler

A Windows app that picks random usernames from an Excel, CSV or text file. Handy for giveaway draws.

## Features

- Three draw modes:
  - **Unique:** no repeats.
  - **With replacement:** repeats allowed.
  - **Elimination:** drawn names are removed from the pool for the next round, until you reset it.
- Quick buttons to pick 1, 5, 10, 25, 50 or all.
- Output options: one per line, space, comma, semicolon or your own separator; add a prefix like `@` or a suffix; numbered list; lowercase / UPPERCASE / Title Case.
- Reads `.xlsx`, `.xlsm`, `.csv` and `.txt`. Skips header rows like "Username" and lets you choose the column.
- Add new usernames (warns about duplicates), search the list, and see the history of draws this session.
- Copy or export the results. Dark and light theme.

## How to use

1. Download the latest ZIP from [Releases](https://github.com/Smokianlord/Username-Shuffler/releases), extract it and run `Username-Shuffler.exe`.
2. It loads a spreadsheet in the same folder (like `Username Dataset.xlsx`), or use **Open File**.
3. Enter how many names to pick (or click a quick button) and draw.
4. Copy with `Ctrl + C`.

From source (Python 3.10+, no extra packages needed):

```bash
python app.py
```

or `pythonw Username-Shuffler.pyw` to run without a console window.

## Keyboard shortcuts

| Shortcut | Action |
| :--- | :--- |
| `Ctrl + O` | Open a file |
| `Ctrl + R` / `F5` | Reload the file |
| `Enter` / `Ctrl + Enter` | Draw |
| `Ctrl + C` | Copy results |
| `Ctrl + E` / `Ctrl + S` | Export results |
| `Ctrl + D` | Dark / light theme |
| `Ctrl + K` | Clear results |

## File formats

- **Excel:** usernames in a column. If there are several columns, pick one in **Active Column**.
- **CSV:** comma, tab or semicolon separated (UTF-8 or Windows encoding).
- **Text:** one username per line.

## Build

```bash
pip install pyinstaller
python build.py
```

The exe and ZIP are created in `dist/`.
