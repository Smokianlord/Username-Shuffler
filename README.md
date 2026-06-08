# Username Shuffler v2.0.0

A clean desktop app for instantly shuffling usernames from an Excel or CSV file.

<p align="center">
  <img width="800" alt="Username Shuffler v2.0.0" src="https://github.com/user-attachments/assets/ddc83302-2534-4fd4-a3e9-014d77f9f6ce" />
</p>

## Features

* Instantly shuffle usernames after typing a number
* Copy shuffled usernames with one click
* Add and save new usernames to Excel/CSV
* Works with `.xlsx`, `.xlsm`, and `.csv` files
* No annoying popup while loading usernames

## How to Use

1. Download the latest release.
2. Extract the ZIP file.
3. Place your username Excel/CSV file in the same folder.
4. Open `Run-App.bat`.
5. Type how many usernames you want.
6. Click **Copy Result** to copy the shuffled usernames.

## Username File Format

Your Excel or CSV file should contain usernames in a column.

Example:

```text
username1
username2
username3
username4
```

The app will automatically detect and load the username list from the file.

## Notes

* Keep your username file in the same folder as the app.
* Duplicate usernames are skipped when saving new usernames.
* The app is made with Python.

## Version

Current version: `v2.0.0`
