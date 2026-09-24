"""
Build and Release Packaging Script for Username Shuffler v3.0.0
Produces standalone Windows executable and release ZIP bundle.
"""
from __future__ import annotations

import hashlib
import os
import shutil
import subprocess
import sys
import zipfile

# Ensure UTF-8 output on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

APP_VERSION = "v3.0.0"
APP_NAME = "Username-Shuffler"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DIST_DIR = os.path.join(BASE_DIR, "dist")
BUILD_DIR = os.path.join(BASE_DIR, "build")
RELEASE_FOLDER_NAME = f"{APP_NAME}-{APP_VERSION}-Windows"
RELEASE_DIR = os.path.join(DIST_DIR, RELEASE_FOLDER_NAME)
RELEASE_ZIP = os.path.join(DIST_DIR, f"{RELEASE_FOLDER_NAME}.zip")


def calculate_sha256(file_path: str) -> str:
    sha = hashlib.sha256()
    with open(file_path, "rb") as f:
        while chunk := f.read(8192):
            sha.update(chunk)
    return sha.hexdigest()


def main() -> int:
    print("==================================================")
    print(f" Building {APP_NAME} {APP_VERSION}")
    print("==================================================")

    # 1. Check PyInstaller
    try:
        import PyInstaller
        print(f"[OK] PyInstaller found (v{PyInstaller.__version__})")
    except ImportError:
        print("[ERROR] PyInstaller is required to build the executable.")
        print("Run: pip install pyinstaller")
        return 1

    # 2. Clean previous build artifacts
    print("\nCleaning previous build artifacts...")
    for path in [DIST_DIR, BUILD_DIR]:
        if os.path.exists(path):
            try:
                shutil.rmtree(path)
                print(f"  Removed {os.path.basename(path)}/")
            except Exception as e:
                print(f"  Warning: Could not remove {path}: {e}")

    # 3. Run PyInstaller
    spec_file = os.path.join(BASE_DIR, "Username-Shuffler.spec")
    print(f"\nRunning PyInstaller with spec: {os.path.basename(spec_file)}...")
    cmd = [sys.executable, "-m", "PyInstaller", spec_file, "--noconfirm"]
    result = subprocess.run(cmd, cwd=BASE_DIR)
    if result.returncode != 0:
        print("\n[ERROR] PyInstaller build failed!")
        return result.returncode

    exe_path = os.path.join(DIST_DIR, f"{APP_NAME}.exe")
    if not os.path.exists(exe_path):
        print(f"\n[ERROR] Expected executable not found at {exe_path}")
        return 1

    exe_size_mb = os.path.getsize(exe_path) / (1024 * 1024)
    print(f"\n[OK] Successfully generated {os.path.basename(exe_path)} ({exe_size_mb:.2f} MB)")

    # Copy EXE directly to root main folder for convenience
    root_exe = os.path.join(BASE_DIR, f"{APP_NAME}.exe")
    shutil.copy2(exe_path, root_exe)
    print(f"[OK] Main folder executable updated: {os.path.basename(root_exe)}")

    # 4. Prepare Release Bundle Folder
    print(f"\nPreparing release package: {RELEASE_FOLDER_NAME}...")
    os.makedirs(RELEASE_DIR, exist_ok=True)

    # Copy EXE to release folder
    shutil.copy2(exe_path, os.path.join(RELEASE_DIR, f"{APP_NAME}.exe"))

    # Copy Sample Dataset
    sample_dataset = os.path.join(BASE_DIR, "Username Dataset.xlsx")
    if not os.path.exists(sample_dataset):
        sample_dataset = os.path.join(BASE_DIR, "Usernames.xlsx")
    if os.path.exists(sample_dataset):
        shutil.copy2(sample_dataset, os.path.join(RELEASE_DIR, os.path.basename(sample_dataset)))
        print(f"  [OK] Included sample '{os.path.basename(sample_dataset)}'")

    # Create Quick Start Guide
    readme_content = f"""Username Shuffler {APP_VERSION}
===================================================

A clean, modern desktop tool for shuffling, formatting, and managing usernames.

QUICK START:
1. Double-click 'Username-Shuffler.exe' to launch.
2. The app automatically detects 'Username Dataset.xlsx' in the folder.
   Or click 'Open File' on the top toolbar to choose any .xlsx, .xlsm, .csv, or .txt file.
3. Enter the number of usernames you want to pick.
4. Click 'Copy Results' or press Ctrl+C to copy.
5. Toggle Dark/Light mode anytime with the toolbar button or press Ctrl+D.

KEYBOARD SHORTCUTS:
• Ctrl + O       : Open Dataset File
• Ctrl + R / F5  : Reload Current File
• Ctrl + Enter   : Shuffle Now
• Ctrl + C       : Copy Formatted Results to Clipboard
• Ctrl + E       : Export Results (TXT, CSV, XLSX)
• Ctrl + D       : Toggle Dark / Light Theme
• Ctrl + K       : Clear Results

GitHub Repository: https://github.com/Smokianlord/Username-Shuffler
"""
    with open(os.path.join(RELEASE_DIR, "README.txt"), "w", encoding="utf-8") as f:
        f.write(readme_content)
    print("  [OK] Created release README.txt")

    # 5. Create ZIP Archive
    print(f"\nCompressing into {os.path.basename(RELEASE_ZIP)}...")
    with zipfile.ZipFile(RELEASE_ZIP, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for root_dir, _, files in os.walk(RELEASE_DIR):
            for file in files:
                abs_path = os.path.join(root_dir, file)
                rel_path = os.path.relpath(abs_path, DIST_DIR)
                zf.write(abs_path, rel_path)

    zip_size_mb = os.path.getsize(RELEASE_ZIP) / (1024 * 1024)
    zip_sha = calculate_sha256(RELEASE_ZIP)
    exe_sha = calculate_sha256(root_exe)

    # Clean up staging directory and loose exe from dist/
    if os.path.exists(RELEASE_DIR):
        shutil.rmtree(RELEASE_DIR)
    if os.path.exists(exe_path):
        try:
            os.remove(exe_path)
        except Exception:
            pass

    print("\n==================================================")
    print(" BUILD SUCCESSFUL!")
    print("==================================================")
    print(f" Executable: {exe_path}")
    print(f" Size:       {exe_size_mb:.2f} MB")
    print(f" SHA256:     {exe_sha}")
    print("--------------------------------------------------")
    print(f" Release ZIP:{RELEASE_ZIP}")
    print(f" Size:       {zip_size_mb:.2f} MB")
    print(f" SHA256:     {zip_sha}")
    print("==================================================\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
