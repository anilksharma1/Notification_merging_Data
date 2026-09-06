# ============================================================
#  File Copier — Search Subfolders & Copy by Excel List
#  - Reads file names from an Excel column
#  - Searches all subfolders of a source folder
#  - COPIES (not moves) found files to a destination folder
#  - Creates a summary report of copied / not found / errors
# ============================================================
#
#  INSTALL: pip install openpyxl
# ============================================================
 
import os
import shutil
import openpyxl
import csv as csv_mod
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
 
# ── Hide root Tkinter window ─────────────────────────────────
root = tk.Tk()
root.withdraw()
root.attributes("-topmost", True)
 
print("=" * 60)
print("       File Copier — Search & Copy by Excel List")
print("=" * 60)
 
# ── Step 1: Select Excel file with file names ─────────────────
print("\n📋 Select EXCEL file containing file names...")
excel_file = filedialog.askopenfilename(
    title="Select Excel File with File Names",
    filetypes=[("Excel Files", "*.xlsx *.xlsm *.xls")]
)
if not excel_file:
    messagebox.showerror("Cancelled", "No Excel file selected. Exiting.")
    exit()
print(f"✅ Excel file : {excel_file}")
 
# ── Step 2: Ask which column has file names ───────────────────
col_input = simpledialog.askstring(
    "Column",
    "Which column has the file names? (e.g. A, B, C)\nLeave blank for A:",
    parent=root
)
col_input = (col_input or "A").strip().upper()
 
# ── Step 3: Ask starting row ──────────────────────────────────
row_input = simpledialog.askstring(
    "Starting Row",
    "Which row do file names start? (e.g. 2 if row 1 is header)\nLeave blank for 2:",
    parent=root
)
try:
    start_row = int(row_input.strip()) if row_input else 2
except ValueError:
    start_row = 2
 
# ── Step 4: Read file names from Excel ───────────────────────
import tempfile as _tmp, shutil as _sh
tmp_excel = _tmp.mktemp(suffix=".xlsx")
_sh.copy2(excel_file, tmp_excel)
wb      = openpyxl.load_workbook(tmp_excel, data_only=True)
ws      = wb.active
col_idx = openpyxl.utils.column_index_from_string(col_input)
 
file_names = []
for row in ws.iter_rows(min_row=start_row, min_col=col_idx, max_col=col_idx, values_only=True):
    val = row[0]
    if val and str(val).strip():
        file_names.append(str(val).strip())
 
try:
    os.remove(tmp_excel)
except:
    pass
 
if not file_names:
    messagebox.showerror("No Data", f"No file names found in column {col_input} starting row {start_row}.")
    exit()
 
print(f"✅ File names loaded : {len(file_names)}")
for f in file_names[:5]:
    print(f"   • {f}")
if len(file_names) > 5:
    print(f"   ... and {len(file_names)-5} more")
 
# ── Step 5: Select source (main) folder ──────────────────────
print("\n📂 Select SOURCE folder (will search all subfolders)...")
source_folder = filedialog.askdirectory(title="Select Source Folder (Search Here)")
if not source_folder:
    messagebox.showerror("Cancelled", "No source folder selected. Exiting.")
    exit()
print(f"✅ Source folder : {source_folder}")
 
# ── Step 6: Select destination folder ────────────────────────
print("\n📁 Select DESTINATION folder (copy files here)...")
dest_folder = filedialog.askdirectory(title="Select Destination Folder (Copy Files Here)")
if not dest_folder:
    messagebox.showerror("Cancelled", "No destination folder selected. Exiting.")
    exit()
print(f"✅ Destination   : {dest_folder}")
 
# ── Step 7: Ask for summary output folder ────────────────────
print("\n📊 Select folder to save SUMMARY report...")
summary_folder = filedialog.askdirectory(title="Select Folder to Save Summary Report")
if not summary_folder:
    summary_folder = dest_folder
summary_path = os.path.join(summary_folder, "copy_summary.csv")
 
# ── Step 8: Ask duplicate handling ───────────────────────────
dup_choice = messagebox.askyesno(
    "Duplicate Files",
    "If a file with the same name already exists in destination,\n"
    "do you want to RENAME it (add _1, _2...)?\n\n"
    "Yes = Rename   |   No = Skip"
)
 
# ── Step 9: Progress Window ───────────────────────────────────
progress_win = tk.Toplevel(root)
progress_win.title("Copying Files...")
progress_win.geometry("620x340")
progress_win.resizable(False, False)
progress_win.attributes("-topmost", True)
 
tk.Label(progress_win, text="File Copier", font=("Segoe UI", 13, "bold")).pack(pady=(14, 4))
tk.Label(progress_win, text="Overall Progress:", font=("Segoe UI", 9)).pack(anchor="w", padx=20)
 
overall_bar = ttk.Progressbar(progress_win, length=575, mode="determinate", maximum=len(file_names))
overall_bar.pack(padx=20, pady=(2, 4))
 
overall_label = tk.Label(
    progress_win,
    text=f"0 of {len(file_names)} files  |  {len(file_names)} pending",
    font=("Segoe UI", 9), fg="#555"
)
overall_label.pack()
 
current_label = tk.Label(progress_win, text="Building file index...",
                          font=("Segoe UI", 9, "italic"), fg="#333", wraplength=600)
current_label.pack(pady=(6, 4))
 
stats_frame = tk.Frame(progress_win, bd=1, relief="groove", padx=10, pady=8)
stats_frame.pack(padx=20, pady=10, fill="x")
 
def stat_col(parent, label, var, col):
    tk.Label(parent, text=label, font=("Segoe UI", 8), fg="#777").grid(row=0, column=col, padx=14)
    tk.Label(parent, textvariable=var, font=("Segoe UI", 12, "bold")).grid(row=1, column=col, padx=14)
 
v_total    = tk.StringVar(value=str(len(file_names)))
v_copied   = tk.StringVar(value="0")
v_notfound = tk.StringVar(value="0")
v_skipped  = tk.StringVar(value="0")
v_errors   = tk.StringVar(value="0")
v_pending  = tk.StringVar(value=str(len(file_names)))
 
stat_col(stats_frame, "Total",       v_total,    0)
stat_col(stats_frame, "✅ Copied",    v_copied,   1)
stat_col(stats_frame, "⏭️ Pending",   v_pending,  2)
stat_col(stats_frame, "❓Not Found",  v_notfound, 3)
stat_col(stats_frame, "⚠️ Skipped",   v_skipped,  4)
stat_col(stats_frame, "❌ Errors",    v_errors,   5)
 
progress_win.update()
 
# ── Step 10: Build file index from source folder ─────────────
current_label.config(text="Scanning all subfolders — building file index...")
progress_win.update()
print("\n🔍 Scanning all subfolders...")
 
file_index = {}
for dirpath, _, filenames in os.walk(source_folder):
    for fname in filenames:
        key = fname.lower()
        full_path = os.path.join(dirpath, fname)
        if key not in file_index:
            file_index[key] = []
        file_index[key].append(full_path)
 
print(f"✅ Indexed {len(file_index)} unique file names across all subfolders")
 
# ── Helper: unique destination path ──────────────────────────
def unique_dest_path(dest_folder, filename):
    dest = os.path.join(dest_folder, filename)
    if not os.path.exists(dest):
        return dest
    name, ext = os.path.splitext(filename)
    counter = 1
    while True:
        new_dest = os.path.join(dest_folder, f"{name}_{counter}{ext}")
        if not os.path.exists(new_dest):
            return new_dest
        counter += 1
 
# ── Step 11: Copy files ───────────────────────────────────────
total_copied   = 0
total_notfound = 0
total_skipped  = 0
total_errors   = 0
summary_rows   = []
 
for idx, file_name in enumerate(file_names, start=1):
    current_label.config(text=f"Copying ({idx}/{len(file_names)}): {file_name}")
    overall_bar["value"] = idx - 1
    overall_label.config(text=f"{idx-1} of {len(file_names)}  |  {len(file_names)-idx+1} pending")
    progress_win.update()
 
    key = file_name.lower()
 
    if key not in file_index:
        print(f"   ❓ Not found : {file_name}")
        total_notfound += 1
        summary_rows.append({
            "File Name"  : file_name,
            "Status"     : "Not Found",
            "Source Path": "",
            "Dest Path"  : "",
            "Note"       : "File not found in any subfolder"
        })
 
    else:
        src_paths = file_index[key]
 
        for i, src_path in enumerate(src_paths):
            if i == 0:
                dest_path = os.path.join(dest_folder, file_name)
 
                if os.path.exists(dest_path):
                    if dup_choice:
                        dest_path = unique_dest_path(dest_folder, file_name)
                        note = f"Renamed to {os.path.basename(dest_path)}"
                    else:
                        print(f"   ⚠️  Skipped (exists) : {file_name}")
                        total_skipped += 1
                        summary_rows.append({
                            "File Name"  : file_name,
                            "Status"     : "Skipped",
                            "Source Path": src_path,
                            "Dest Path"  : dest_path,
                            "Note"       : "Already exists in destination — skipped"
                        })
                        continue
                else:
                    note = ""
 
                try:
                    shutil.copy2(src_path, dest_path)   # copy2 preserves metadata
                    total_copied += 1
                    print(f"   ✅ Copied : {file_name}")
                    summary_rows.append({
                        "File Name"  : file_name,
                        "Status"     : "Copied",
                        "Source Path": src_path,
                        "Dest Path"  : dest_path,
                        "Note"       : note
                    })
                except Exception as e:
                    total_errors += 1
                    print(f"   ❌ Error : {file_name} — {e}")
                    summary_rows.append({
                        "File Name"  : file_name,
                        "Status"     : "Error",
                        "Source Path": src_path,
                        "Dest Path"  : "",
                        "Note"       : str(e)
                    })
            else:
                summary_rows.append({
                    "File Name"  : file_name,
                    "Status"     : "Duplicate in Source",
                    "Source Path": src_path,
                    "Dest Path"  : "",
                    "Note"       : "Additional copy found in source — not copied"
                })
 
    # Update live stats
    v_copied.set(str(total_copied))
    v_notfound.set(str(total_notfound))
    v_skipped.set(str(total_skipped))
    v_errors.set(str(total_errors))
    v_pending.set(str(len(file_names) - idx))
    overall_bar["value"] = idx
    overall_label.config(text=f"{idx} of {len(file_names)}  |  {len(file_names)-idx} pending")
    progress_win.update()
 
# ── Step 12: Write summary CSV ────────────────────────────────
with open(summary_path, "w", newline="", encoding="utf-8-sig") as sf:
    writer = csv_mod.DictWriter(sf, fieldnames=["File Name", "Status", "Source Path", "Dest Path", "Note"])
    writer.writeheader()
    writer.writerows(summary_rows)
    writer.writerow({})
    writer.writerow({"File Name": "── SUMMARY ──"})
    writer.writerow({"File Name": "Total Files in List",  "Status": len(file_names)})
    writer.writerow({"File Name": "Copied Successfully",  "Status": total_copied})
    writer.writerow({"File Name": "Not Found",            "Status": total_notfound})
    writer.writerow({"File Name": "Skipped (Duplicate)",  "Status": total_skipped})
    writer.writerow({"File Name": "Errors",               "Status": total_errors})
 
# ── Done ─────────────────────────────────────────────────────
progress_win.destroy()
 
summary_msg = (
    f"✅ File Copy Complete!\n\n"
    f"📋 Files in List       : {len(file_names)}\n"
    f"✅ Copied Successfully : {total_copied}\n"
    f"❓ Not Found           : {total_notfound}\n"
    f"⚠️  Skipped             : {total_skipped}\n"
    f"❌ Errors              : {total_errors}\n\n"
    f"📊 Summary Report      : {summary_path}"
)
 
print("\n" + "=" * 60)
print(summary_msg)
print("=" * 60)
 
messagebox.showinfo("Complete ✅", summary_msg)
root.destroy()
