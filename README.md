# Username Shuffler v2.0.0

A clean desktop app for instantly shuffling usernames from an Excel or CSV file.

## Run

Put your username Excel/CSV file in the same folder as the app files, then double-click:

```text
Run-App.bat
```

## Notes

- The zip is flat now. There is no extra folder inside the zip.
- The copy result area stays visible in the default window size.
- The title-bar icon uses `titlebar.ico`, and the EXE build uses `icon.ico`.
- `.xlsx`, `.xlsm`, and `.csv` loading works without openpyxl.
- Saving new usernames creates or updates `usernames.xlsx` when needed.

## Build EXE on Windows

Double-click:

```text
build_windows.bat
```

The EXE will be created in the `dist` folder.
