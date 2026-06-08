import csv
import os
import random
import re
import sys
import threading
import tkinter as tk
from tkinter import ttk

APP_NAME = "Username Shuffler"
APP_VERSION = "v2.0.0"
DATA_EXTENSIONS = (".xlsx", ".xlsm", ".csv", ".xls")

BG = "#EEF4FF"
HEADER = "#0F172A"
CARD = "#FFFFFF"
CARD_BORDER = "#D8E2F1"
SHADOW = "#C7D2FE"
TEXT = "#111827"
MUTED = "#64748B"
PRIMARY = "#2563EB"
PRIMARY_HOVER = "#1D4ED8"
GREEN = "#16A34A"
GREEN_HOVER = "#15803D"
CYAN = "#0891B2"
CYAN_HOVER = "#0E7490"
ORANGE = "#EA580C"
ORANGE_HOVER = "#C2410C"
RED = "#DC2626"


def resource_path(relative_path: str) -> str:
    """Return resource path for normal Python and PyInstaller builds."""
    try:
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    except Exception:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)


def get_app_folder() -> str:
    """Return the folder where the exe or py file is located."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def clean_username(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "null"}:
        return ""
    return text


def find_data_file(folder: str) -> str | None:
    files = []
    for file_name in os.listdir(folder):
        lower = file_name.lower()
        if file_name.startswith("~$") or file_name.startswith("."):
            continue
        if lower.endswith(DATA_EXTENSIONS):
            files.append(file_name)

    if not files:
        return None

    priority = ["usernames.xlsx", "usernames.csv", "username.xlsx", "username.csv"]
    for preferred in priority:
        for file_name in files:
            if file_name.lower() == preferred:
                return os.path.join(folder, file_name)

    files.sort(key=str.lower)
    return os.path.join(folder, files[0])


def read_csv_file(path: str) -> list[str]:
    usernames: list[str] = []
    encodings = ("utf-8-sig", "utf-8", "cp1252")
    last_error: Exception | None = None

    for encoding in encodings:
        try:
            with open(path, "r", encoding=encoding, newline="") as handle:
                reader = csv.reader(handle)
                for row in reader:
                    for cell in row:
                        name = clean_username(cell)
                        if name:
                            usernames.append(name)
            return usernames
        except UnicodeDecodeError as exc:
            last_error = exc

    if last_error:
        raise last_error
    return usernames


def read_excel_file(path: str) -> list[str]:
    lower = path.lower()

    if lower.endswith((".xlsx", ".xlsm")):
        from openpyxl import load_workbook

        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            sheet = workbook.active
            usernames: list[str] = []
            for row in sheet.iter_rows(values_only=True):
                for cell in row:
                    name = clean_username(cell)
                    if name:
                        usernames.append(name)
            return usernames
        finally:
            workbook.close()

    if lower.endswith(".xls"):
        try:
            import pandas as pd  # Optional fallback for old Excel files.
        except Exception as exc:
            raise RuntimeError(
                "Old .xls files need pandas/xlrd. Please save the sheet as .xlsx or .csv."
            ) from exc

        frame = pd.read_excel(path, header=None)
        return [
            name
            for name in (clean_username(value) for value in frame.values.flatten().tolist())
            if name
        ]

    raise RuntimeError("Unsupported file type.")


def load_usernames_from_folder(folder: str) -> tuple[list[str], str | None, str | None]:
    path = find_data_file(folder)
    if not path:
        return [], None, "No Excel/CSV file found. Add usernames below and save to create usernames.xlsx."

    lower = path.lower()
    if lower.endswith(".csv"):
        usernames = read_csv_file(path)
    else:
        usernames = read_excel_file(path)

    return usernames, path, None


def parse_new_usernames(raw_text: str) -> list[str]:
    parts: list[str] = []
    for line in raw_text.splitlines():
        line = line.strip()
        if not line:
            continue
        # Each line can be one username, or a comma/semicolon separated batch.
        parts.extend(re.split(r"[,;]+", line))
    return [name for name in (clean_username(part) for part in parts) if name]


def append_usernames_to_file(path: str, names: list[str]) -> str:
    lower = path.lower()

    if lower.endswith(".csv"):
        with open(path, "a", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle)
            for name in names:
                writer.writerow([name])
        return path

    if lower.endswith((".xlsx", ".xlsm")):
        from openpyxl import Workbook, load_workbook

        if os.path.exists(path):
            workbook = load_workbook(path)
            sheet = workbook.active
        else:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Usernames"

        for name in names:
            sheet.append([name])

        workbook.save(path)
        workbook.close()
        return path

    # Do not write back to .xls. Create a modern sheet instead.
    new_path = os.path.join(get_app_folder(), "usernames.xlsx")
    return append_usernames_to_file(new_path, names)


class UsernameShufflerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.folder = get_app_folder()
        self.usernames: list[str] = []
        self.loaded_file: str | None = None
        self.result_text = ""
        self.load_error: str | None = None
        self.count_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Starting...")
        self.file_var = tk.StringVar(value="Looking for sheet...")
        self.total_var = tk.StringVar(value="0")
        self.pick_var = tk.StringVar(value="0")
        self._trace_ready = False

        self.configure_window()
        self.build_ui()
        self.count_var.trace_add("write", self.on_count_changed)
        self._trace_ready = True
        self.load_data_async(initial=True)

    def configure_window(self) -> None:
        self.root.title(f"{APP_NAME} {APP_VERSION}")
        self.root.geometry("860x650")
        self.root.minsize(780, 600)
        self.root.configure(bg=BG)

        icon_ico = resource_path("icon.ico")
        icon_png = resource_path("icon.png")
        try:
            if os.path.exists(icon_ico):
                self.root.iconbitmap(icon_ico)
        except Exception:
            pass
        try:
            if os.path.exists(icon_png):
                self.window_icon = tk.PhotoImage(file=icon_png)
                self.root.iconphoto(True, self.window_icon)
        except Exception:
            pass

    def build_ui(self) -> None:
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        self.build_header()

        body = tk.Frame(self.root, bg=BG)
        body.grid(row=1, column=0, sticky="nsew", padx=22, pady=18)
        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(1, weight=1)

        self.build_stats_card(body)
        self.build_shuffle_card(body)
        self.build_result_card(body)
        self.build_add_card(body)

        status_bar = tk.Frame(self.root, bg="#DBEAFE", height=34)
        status_bar.grid(row=2, column=0, sticky="ew")
        status_bar.grid_columnconfigure(0, weight=1)
        tk.Label(
            status_bar,
            textvariable=self.status_var,
            bg="#DBEAFE",
            fg="#1E3A8A",
            font=("Segoe UI", 10, "bold"),
            anchor="w",
            padx=18,
        ).grid(row=0, column=0, sticky="ew", ipady=7)

    def build_header(self) -> None:
        header = tk.Frame(self.root, bg=HEADER, height=105)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(1, weight=1)
        header.grid_propagate(False)

        logo_holder = tk.Frame(header, bg=HEADER)
        logo_holder.grid(row=0, column=0, padx=(24, 14), pady=22, sticky="w")
        try:
            self.logo_image = tk.PhotoImage(file=resource_path("icon.png"))
            self.logo_image = self.logo_image.subsample(max(1, self.logo_image.width() // 56))
            tk.Label(logo_holder, image=self.logo_image, bg=HEADER).pack()
        except Exception:
            tk.Label(
                logo_holder,
                text="US",
                bg=PRIMARY,
                fg="white",
                font=("Segoe UI", 18, "bold"),
                width=3,
                height=1,
            ).pack()

        title_area = tk.Frame(header, bg=HEADER)
        title_area.grid(row=0, column=1, sticky="w", pady=18)
        tk.Label(
            title_area,
            text=f"{APP_NAME} {APP_VERSION}",
            bg=HEADER,
            fg="white",
            font=("Segoe UI", 22, "bold"),
        ).pack(anchor="w")
        tk.Label(
            title_area,
            text="Instant random picks from your Excel or CSV username list",
            bg=HEADER,
            fg="#CBD5E1",
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(4, 0))

        status_pill = tk.Label(
            header,
            text="Auto Shuffle ON",
            bg="#DCFCE7",
            fg="#166534",
            font=("Segoe UI", 10, "bold"),
            padx=12,
            pady=6,
        )
        status_pill.grid(row=0, column=2, padx=24, sticky="e")

    def card(self, parent: tk.Widget, row: int, column: int, *, columnspan: int = 1, sticky: str = "nsew") -> tk.Frame:
        shadow = tk.Frame(parent, bg=SHADOW)
        shadow.grid(row=row, column=column, columnspan=columnspan, sticky=sticky, padx=8, pady=8)
        shadow.grid_columnconfigure(0, weight=1)
        shadow.grid_rowconfigure(0, weight=1)
        inner = tk.Frame(shadow, bg=CARD, highlightbackground=CARD_BORDER, highlightthickness=1)
        inner.grid(row=0, column=0, sticky="nsew", padx=(0, 4), pady=(0, 4))
        return inner

    def build_stats_card(self, parent: tk.Widget) -> None:
        card = self.card(parent, 0, 0)
        card.grid_columnconfigure(1, weight=1)

        tk.Label(card, text="Dataset", bg=CARD, fg=TEXT, font=("Segoe UI", 14, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=18, pady=(16, 10)
        )
        self.metric_row(card, 1, "Current file", self.file_var)
        self.metric_row(card, 2, "Loaded usernames", self.total_var)
        self.metric_row(card, 3, "Current result", self.pick_var)

    def metric_row(self, parent: tk.Widget, row: int, label: str, value_var: tk.StringVar) -> None:
        tk.Label(parent, text=label, bg=CARD, fg=MUTED, font=("Segoe UI", 10)).grid(
            row=row, column=0, sticky="w", padx=18, pady=5
        )
        tk.Label(parent, textvariable=value_var, bg=CARD, fg=TEXT, font=("Segoe UI", 10, "bold"), anchor="e").grid(
            row=row, column=1, sticky="ew", padx=18, pady=5
        )

    def build_shuffle_card(self, parent: tk.Widget) -> None:
        card = self.card(parent, 0, 1)
        card.grid_columnconfigure(0, weight=1)
        card.grid_columnconfigure(1, weight=1)

        tk.Label(card, text="Shuffle Control", bg=CARD, fg=TEXT, font=("Segoe UI", 14, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=18, pady=(16, 8)
        )
        tk.Label(card, text="Type a number. Result updates instantly.", bg=CARD, fg=MUTED, font=("Segoe UI", 10)).grid(
            row=1, column=0, columnspan=2, sticky="w", padx=18, pady=(0, 10)
        )

        entry = tk.Entry(
            card,
            textvariable=self.count_var,
            justify="center",
            bg="#F8FAFC",
            fg=TEXT,
            insertbackground=TEXT,
            font=("Segoe UI", 24, "bold"),
            relief="solid",
            bd=1,
            highlightthickness=2,
            highlightbackground="#BFDBFE",
            highlightcolor=PRIMARY,
        )
        entry.grid(row=2, column=0, columnspan=2, sticky="ew", padx=18, pady=(0, 14), ipady=8)
        self.count_entry = entry

        self.button(card, "Re-shuffle", PRIMARY, PRIMARY_HOVER, self.shuffle_manual).grid(
            row=3, column=0, sticky="ew", padx=(18, 6), pady=(0, 16), ipady=6
        )
        self.button(card, "Reload Sheet", GREEN, GREEN_HOVER, lambda: self.load_data_async(initial=False)).grid(
            row=3, column=1, sticky="ew", padx=(6, 18), pady=(0, 16), ipady=6
        )

    def build_result_card(self, parent: tk.Widget) -> None:
        card = self.card(parent, 1, 0, columnspan=2)
        card.grid_rowconfigure(1, weight=1)
        card.grid_columnconfigure(0, weight=1)

        top = tk.Frame(card, bg=CARD)
        top.grid(row=0, column=0, sticky="ew", padx=18, pady=(16, 8))
        top.grid_columnconfigure(0, weight=1)
        tk.Label(top, text="Result", bg=CARD, fg=TEXT, font=("Segoe UI", 14, "bold")).grid(row=0, column=0, sticky="w")
        self.button(top, "Copy Result", CYAN, CYAN_HOVER, self.copy_result, width=14).grid(row=0, column=1, sticky="e")

        text_frame = tk.Frame(card, bg="#E2E8F0")
        text_frame.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 16))
        text_frame.grid_columnconfigure(0, weight=1)
        text_frame.grid_rowconfigure(0, weight=1)

        self.output_box = tk.Text(
            text_frame,
            height=5,
            wrap="word",
            bg="#F8FAFC",
            fg=TEXT,
            relief="flat",
            font=("Consolas", 12),
            padx=12,
            pady=12,
        )
        self.output_box.grid(row=0, column=0, sticky="nsew", padx=(0, 1), pady=(0, 1))
        self.output_box.insert("1.0", "Type a number above to generate shuffled usernames instantly.")
        self.output_box.configure(state="disabled")

    def build_add_card(self, parent: tk.Widget) -> None:
        card = self.card(parent, 2, 0, columnspan=2)
        card.grid_columnconfigure(0, weight=1)

        tk.Label(card, text="Add New Usernames", bg=CARD, fg=TEXT, font=("Segoe UI", 14, "bold")).grid(
            row=0, column=0, sticky="w", padx=18, pady=(16, 4)
        )
        tk.Label(
            card,
            text="Paste one username per line. Comma or semicolon separated batches also work.",
            bg=CARD,
            fg=MUTED,
            font=("Segoe UI", 10),
        ).grid(row=1, column=0, sticky="w", padx=18, pady=(0, 8))

        self.add_box = tk.Text(
            card,
            height=4,
            wrap="word",
            bg="#F8FAFC",
            fg=TEXT,
            relief="solid",
            bd=1,
            font=("Segoe UI", 10),
            padx=10,
            pady=8,
        )
        self.add_box.grid(row=2, column=0, sticky="ew", padx=18, pady=(0, 10))

        actions = tk.Frame(card, bg=CARD)
        actions.grid(row=3, column=0, sticky="ew", padx=18, pady=(0, 16))
        actions.grid_columnconfigure(0, weight=1)
        actions.grid_columnconfigure(1, weight=1)
        self.button(actions, "Save to Excel/CSV", ORANGE, ORANGE_HOVER, self.save_new_usernames).grid(
            row=0, column=0, sticky="ew", padx=(0, 6), ipady=6
        )
        self.button(actions, "Clear Box", "#475569", "#334155", self.clear_add_box).grid(
            row=0, column=1, sticky="ew", padx=(6, 0), ipady=6
        )

    def button(self, parent: tk.Widget, text: str, bg: str, hover: str, command, width: int | None = None) -> tk.Button:
        button = tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg="white",
            activebackground=hover,
            activeforeground="white",
            relief="raised",
            bd=3,
            cursor="hand2",
            font=("Segoe UI", 10, "bold"),
            width=width or 0,
            highlightthickness=0,
        )
        button.bind("<Enter>", lambda event: button.configure(bg=hover))
        button.bind("<Leave>", lambda event: button.configure(bg=bg))
        return button

    def set_status(self, message: str, *, error: bool = False) -> None:
        self.status_var.set(message)
        if error:
            # Keep the app non-intrusive: errors stay in the status bar.
            self.root.bell()

    def set_output(self, text: str) -> None:
        self.output_box.configure(state="normal")
        self.output_box.delete("1.0", tk.END)
        self.output_box.insert("1.0", text)
        self.output_box.configure(state="disabled")

    def update_metrics(self) -> None:
        self.total_var.set(str(len(self.usernames)))
        self.pick_var.set(str(len(self.result_text.split())) if self.result_text else "0")
        if self.loaded_file:
            self.file_var.set(os.path.basename(self.loaded_file))
        else:
            self.file_var.set("No sheet yet")

    def load_data_async(self, *, initial: bool) -> None:
        self.set_status("Loading usernames from the app folder...")
        self.file_var.set("Loading...")

        def worker() -> None:
            try:
                names, path, warning = load_usernames_from_folder(self.folder)
                error = None
            except Exception as exc:
                names, path, warning, error = [], None, None, str(exc)
            self.root.after(0, lambda: self.finish_load(names, path, warning, error, initial))

        threading.Thread(target=worker, daemon=True).start()

    def finish_load(
        self,
        names: list[str],
        path: str | None,
        warning: str | None,
        error: str | None,
        initial: bool,
    ) -> None:
        self.usernames = names
        self.loaded_file = path
        self.result_text = ""
        self.update_metrics()

        if error:
            self.set_output("Could not load usernames. Check the status message below.")
            self.set_status(f"Load error: {error}", error=True)
        elif warning:
            self.set_output("No usernames loaded yet. Add usernames below and save them to start.")
            self.set_status(warning)
        else:
            self.set_status(f"Loaded {len(names)} usernames from {os.path.basename(path or '')}.")
            if self.count_var.get().strip():
                self.shuffle_names(automatic=True)
            else:
                self.set_output("Type a number above to generate shuffled usernames instantly.")

        if initial:
            self.count_entry.focus_force()

    def on_count_changed(self, *_args) -> None:
        if not self._trace_ready:
            return
        self.shuffle_names(automatic=True)

    def shuffle_manual(self) -> None:
        self.shuffle_names(automatic=False)

    def shuffle_names(self, *, automatic: bool) -> None:
        raw_value = self.count_var.get().strip()

        if not raw_value:
            self.result_text = ""
            self.pick_var.set("0")
            self.set_output("Type a number above to generate shuffled usernames instantly.")
            if self.usernames:
                self.set_status(f"Ready. {len(self.usernames)} usernames loaded.")
            return

        if not raw_value.isdigit():
            self.result_text = ""
            self.pick_var.set("0")
            self.set_output("Only numbers are accepted here.")
            self.set_status("Enter a valid positive number.", error=not automatic)
            return

        amount = int(raw_value)
        if amount <= 0:
            self.result_text = ""
            self.pick_var.set("0")
            self.set_output("Number must be greater than 0.")
            self.set_status("Enter a number greater than 0.", error=not automatic)
            return

        if not self.usernames:
            self.result_text = ""
            self.pick_var.set("0")
            self.set_output("No usernames available yet. Add usernames below and save them first.")
            self.set_status("No usernames loaded. Save usernames to create a sheet.", error=not automatic)
            return

        if amount > len(self.usernames):
            self.result_text = ""
            self.pick_var.set("0")
            self.set_output(f"You asked for {amount}, but only {len(self.usernames)} usernames are loaded.")
            self.set_status("Pick count cannot be larger than loaded usernames.", error=not automatic)
            return

        selected = random.sample(self.usernames, amount)
        self.result_text = " ".join(selected)
        self.pick_var.set(str(amount))
        self.set_output(self.result_text)
        self.set_status(f"Generated {amount} shuffled username{'s' if amount != 1 else ''}.")

    def copy_result(self) -> None:
        if not self.result_text:
            self.set_status("Nothing to copy yet. Enter a number first.", error=True)
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(self.result_text)
        self.root.update_idletasks()
        self.set_status("Result copied to clipboard.")
        self.count_entry.focus_force()

    def clear_add_box(self) -> None:
        self.add_box.delete("1.0", tk.END)
        self.set_status("Add box cleared.")

    def save_new_usernames(self) -> None:
        raw_text = self.add_box.get("1.0", tk.END)
        parsed = parse_new_usernames(raw_text)

        if not parsed:
            self.set_status("No new usernames to save.", error=True)
            return

        existing = {name.casefold() for name in self.usernames}
        seen_new: set[str] = set()
        unique_new: list[str] = []
        skipped = 0

        for name in parsed:
            key = name.casefold()
            if key in existing or key in seen_new:
                skipped += 1
                continue
            seen_new.add(key)
            unique_new.append(name)

        if not unique_new:
            self.set_status(f"No new usernames saved. {skipped} duplicate item(s) skipped.", error=True)
            return

        target_path = self.loaded_file
        if not target_path or target_path.lower().endswith(".xls"):
            target_path = os.path.join(self.folder, "usernames.xlsx")

        self.set_status("Saving new usernames...")

        def worker() -> None:
            try:
                saved_path = append_usernames_to_file(target_path, unique_new)
                error = None
            except Exception as exc:
                saved_path = target_path
                error = str(exc)
            self.root.after(0, lambda: self.finish_save(unique_new, skipped, saved_path, error))

        threading.Thread(target=worker, daemon=True).start()

    def finish_save(self, added: list[str], skipped: int, saved_path: str, error: str | None) -> None:
        if error:
            self.set_status(f"Save error: {error}", error=True)
            return

        self.loaded_file = saved_path
        self.usernames.extend(added)
        self.add_box.delete("1.0", tk.END)
        self.update_metrics()
        message = f"Saved {len(added)} new username{'s' if len(added) != 1 else ''} to {os.path.basename(saved_path)}."
        if skipped:
            message += f" Skipped {skipped} duplicate item(s)."
        self.set_status(message)
        if self.count_var.get().strip():
            self.shuffle_names(automatic=True)


def main() -> None:
    root = tk.Tk()
    UsernameShufflerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
