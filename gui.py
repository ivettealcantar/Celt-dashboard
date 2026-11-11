import asyncio
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext

import pandas as pd

import bill_checker

log_queue = queue.Queue()


class QueueWriter:
    def write(self, text):
        if text:
            log_queue.put(text)

    def flush(self):
        pass


root = tk.Tk()
root.title("National Grid Bill Checker")
root.geometry("820x580")

input_var = tk.StringVar()
output_var = tk.StringVar()
headless_var = tk.BooleanVar(value=False)
status_var = tk.StringVar(value="Ready")


def gui_pause(message):
    """Called from the worker thread; shows a dialog and waits for OK."""
    done = threading.Event()

    def show():
        messagebox.showinfo("Action needed", message.strip())
        done.set()

    root.after(0, show)
    done.wait()
    return ""


bill_checker.pause = gui_pause


def open_file(path):
    if sys.platform == "darwin":
        subprocess.run(["open", str(path)])
    elif sys.platform.startswith("win"):
        os.startfile(str(path))
    else:
        subprocess.run(["xdg-open", str(path)])


def pick_input():
    path = filedialog.askopenfilename(
        filetypes=[("Excel or CSV files", "*.xlsx *.xls *.csv")]
    )
    if path:
        input_var.set(path)
        if not output_var.get():
            output_var.set(
                str(Path(path).with_name(Path(path).stem + "_report.xlsx"))
            )


def pick_output():
    path = filedialog.asksaveasfilename(
        defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")]
    )
    if path:
        output_var.set(path)


def make_template():
    path = filedialog.asksaveasfilename(
        defaultextension=".xlsx", initialfile="accounts_template.xlsx"
    )
    if path:
        pd.DataFrame(
            {"Account Number": ["0123456789"], "Account Address": [""]}
        ).to_excel(path, index=False)
        messagebox.showinfo("Saved", "Template saved.")


def check_file():
    if not input_var.get():
        messagebox.showwarning("Missing", "Choose an accounts file first.")
        return None
    try:
        bill_checker.INPUT_FILE = Path(input_var.get())
        accounts = bill_checker.load_accounts()
    except Exception as exc:
        messagebox.showerror("Problem with file", str(exc))
        return None

    preview = "\n".join(accounts["Account Number"].head(5))
    more = "\n..." if len(accounts) > 5 else ""
    messagebox.showinfo(
        "File looks good",
        f"Found {len(accounts)} account(s):\n\n{preview}{more}",
    )
    return accounts


def finished(results, error):
    start_btn.config(state="normal")
    root.bell()  # system beep
    root.lift()  # bring the window to the front
    root.attributes("-topmost", True)
    root.after(500, lambda: root.attributes("-topmost", False))

    if error:
        status_var.set("Stopped with an error")
        messagebox.showerror("Stopped", f"The run stopped:\n\n{error}")
        return

    total = len(results or [])
    needs_review = sum(
        1 for r in (results or []) if r.get("Result") == "Needs review"
    )
    status_var.set(f"Finished: {total} rows, {needs_review} need review")

    if messagebox.askyesno(
        "Finished",
        f"Done! {total} bill rows saved.\n"
        f"{needs_review} row(s) marked 'Needs review'.\n\n"
        f"Open the report now?",
    ):
        open_file(bill_checker.OUTPUT_FILE)


def run_job():
    sys.stdout = QueueWriter()
    sys.stderr = sys.stdout
    results, error = None, None
    try:
        results = asyncio.run(bill_checker.main())
    except Exception as exc:
        error = exc
        print(f"\nERROR: {exc}")
    finally:
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__
        root.after(0, lambda: finished(results, error))


def start():
    if not input_var.get() or not output_var.get():
        messagebox.showwarning("Missing", "Choose an input and output file.")
        return

    if check_file() is None:
        return

    bill_checker.INPUT_FILE = Path(input_var.get())
    bill_checker.OUTPUT_FILE = Path(output_var.get())
    # Hidden mode only works once a saved login exists.
    bill_checker.HEADLESS = (
        headless_var.get() and bill_checker.AUTH_FILE.exists()
    )

    start_btn.config(state="disabled")
    status_var.set("Running... do not close the browser")
    log.delete("1.0", tk.END)
    threading.Thread(target=run_job, daemon=True).start()


def poll_log():
    while not log_queue.empty():
        log.insert(tk.END, log_queue.get())
        log.see(tk.END)
    root.after(100, poll_log)


frame = tk.Frame(root, padx=12, pady=12)
frame.pack(fill="x")

tk.Label(frame, text="Accounts file:").grid(row=0, column=0, sticky="w")
tk.Entry(frame, textvariable=input_var, width=65).grid(row=0, column=1)
tk.Button(frame, text="Browse", command=pick_input).grid(row=0, column=2)
tk.Button(frame, text="Check file", command=check_file).grid(row=0, column=3)

tk.Label(frame, text="Save report as:").grid(row=1, column=0, sticky="w")
tk.Entry(frame, textvariable=output_var, width=65).grid(row=1, column=1)
tk.Button(frame, text="Browse", command=pick_output).grid(row=1, column=2)

tk.Checkbutton(
    frame,
    text="Run browser hidden (after first login)",
    variable=headless_var,
).grid(row=2, column=1, sticky="w")

start_btn = tk.Button(frame, text="Start", command=start, width=14)
start_btn.grid(row=3, column=1, sticky="w", pady=8)
tk.Button(
    frame, text="Download input template", command=make_template
).grid(row=3, column=1, sticky="e")

log = scrolledtext.ScrolledText(root, height=20)
log.pack(fill="both", expand=True, padx=12, pady=(0, 6))

tk.Label(root, textvariable=status_var, anchor="w").pack(
    fill="x", padx=12, pady=(0, 8)
)

poll_log()
root.mainloop()
