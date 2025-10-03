#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Oct  2 18:34:01 2025

@author: mr.basix
"""

# Venn_app.py
# Interactive multi-sheet Venn app:
# - Load workbook → browse sheets with Prev/Next
# - Rename labels, pick colors, set alpha & label height
# - Symmetric diagram (counts), embedded in window
# - Save PNG+CSV per sheet, Save All, optional combined xlsx

import tkinter as tk
from tkinter import filedialog, colorchooser, messagebox, scrolledtext, ttk
import pandas as pd
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.patches import Circle
import os, re, datetime as dt

# ---------- helpers ----------
def sanitize_filename(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"[^\w\-. ]+", "_", s)
    return s[:60] if len(s) > 60 else s

def compute_sets(seriesA, seriesB):
    setA = set(seriesA.dropna().astype(str))
    setB = set(seriesB.dropna().astype(str))
    unique_A = sorted(setA - setB)
    unique_B = sorted(setB - setA)
    shared   = sorted(setA & setB)
    return setA, setB, unique_A, unique_B, shared

def draw_symmetric(ax, setA, setB, la, lb, colorA, colorB, alpha, label_y):
    # clear & draw
    ax.clear()
    r = 1.5
    cx1, cy1 = -0.9, 0
    cx2, cy2 =  0.9, 0

    # circles
    ax.add_patch(Circle((cx1, cy1), r, alpha=alpha, facecolor=colorA, edgecolor="black"))
    ax.add_patch(Circle((cx2, cy2), r, alpha=alpha, facecolor=colorB, edgecolor="black"))

    unique_A = len(setA - setB)
    unique_B = len(setB - setA)
    shared   = len(setA & setB)

    # labels above circles
    ax.text(cx1, r*label_y, la, ha="center", va="bottom", fontsize=13)
    ax.text(cx2, r*label_y, lb, ha="center", va="bottom", fontsize=13)

    # counts
    ax.text(cx1 - 0.4, 0, str(unique_A), ha="center", va="center", fontsize=16)
    ax.text(cx2 + 0.4, 0, str(unique_B), ha="center", va="center", fontsize=16)
    ax.text(0, 0, str(shared), ha="center", va="center", fontsize=16, fontweight="bold")

    ax.set_aspect("equal"); ax.set_xlim(-3,3); ax.set_ylim(-2.5,2.5)
    ax.axis("off")
    ax.set_title("Symmetric Venn Diagram (Counts)")

def to_results_df(la, lb, unique_A, unique_B, shared):
    max_len = max(len(unique_A), len(unique_B), len(shared))
    pad = lambda L: L + [""]*(max_len - len(L))
    return pd.DataFrame({
        f"Unique to {la}": pad(unique_A),
        f"Unique to {lb}": pad(unique_B),
        "Shared": pad(shared)
    })

# ---------- app state ----------
class AppState:
    def __init__(self):
        self.wb_path = None
        self.sheets = []          # list[str]
        self.dfs = {}             # name -> DataFrame
        self.labels = {}          # name -> (labelA, labelB) editable
        self.headers = {}         # name -> (orig_headerA, orig_headerB)
        self.idx = 0

        # visuals
        self.colorA = "#f4c27a"
        self.colorB = "#a6d49f"
        self.alpha  = 0.45
        self.label_y = 1.12

    def has_data(self):
        return bool(self.dfs)

S = AppState()

# ---------- UI callbacks ----------
def load_workbook():
    path = filedialog.askopenfilename(title="Select Excel Workbook",
                                      filetypes=[("Excel files","*.xlsx *.xls")])
    if not path:
        return
    try:
        wb = pd.ExcelFile(path)
        S.wb_path = path
        S.sheets = []
        S.dfs.clear()
        S.labels.clear()
        S.headers.clear()

        for name in wb.sheet_names:
            df = wb.parse(name)
            if df.shape[1] >= 2:
                S.sheets.append(name)
                S.dfs[name] = df
                hdrA, hdrB = str(df.columns[0]), str(df.columns[1])
                S.headers[name] = (hdrA, hdrB)
                S.labels[name]  = [hdrA, hdrB]  # start editable with headers

        if not S.sheets:
            messagebox.showerror("No usable sheets",
                                 "No sheets with at least two columns (A & B) found.")
            return

        S.idx = 0
        refresh_sheet_ui()
        messagebox.showinfo("Loaded",
                            f"Loaded {os.path.basename(path)}\n"
                            f"{len(S.sheets)} sheet(s) ready.")
    except Exception as e:
        messagebox.showerror("Error loading workbook", str(e))

def refresh_sheet_ui():
    if not S.has_data(): return
    name = S.sheets[S.idx]
    df = S.dfs[name]
    la, lb = S.labels[name]

    lbl_file.config(text=f"Workbook: {os.path.basename(S.wb_path) if S.wb_path else ''}")
    lbl_sheet.config(text=f"Sheet [{S.idx+1}/{len(S.sheets)}]: {name}")

    entry_labelA.delete(0, tk.END); entry_labelA.insert(0, la)
    entry_labelB.delete(0, tk.END); entry_labelB.insert(0, lb)

    # preview counts
    setA, setB, uA, uB, sh = compute_sets(df.iloc[:,0], df.iloc[:,1])
    lbl_preview.config(text=f"Preview — Unique A: {len(uA)} | Shared: {len(sh)} | Unique B: {len(uB)}")

    # redraw plot
    draw_symmetric(ax, setA, setB, la, lb, S.colorA, S.colorB, S.alpha, S.label_y)
    canvas.draw()

    # show lists in panel (trim super long to keep UI snappy)
    def preview_list(name_tag, L):
        head = L[:50]
        extra = f"\n…(+{len(L)-50} more)" if len(L) > 50 else ""
        return f"{name_tag} ({len(L)}):\n" + "\n".join(head) + extra + "\n\n"
    output_box.config(state="normal"); output_box.delete(1.0, tk.END)
    output_box.insert(tk.END, preview_list(f"Unique to {la}", uA))
    output_box.insert(tk.END, preview_list(f"Unique to {lb}", uB))
    output_box.insert(tk.END, preview_list("Shared", sh))
    output_box.config(state="disabled")

def apply_label_changes():
    if not S.has_data(): return
    name = S.sheets[S.idx]
    la = entry_labelA.get().strip() or S.headers[name][0]
    lb = entry_labelB.get().strip() or S.headers[name][1]
    S.labels[name] = [la, lb]
    refresh_sheet_ui()

def reset_labels_to_headers():
    if not S.has_data(): return
    name = S.sheets[S.idx]
    S.labels[name] = list(S.headers[name])
    refresh_sheet_ui()

def pick_colorA():
    c = colorchooser.askcolor(title="Pick color for Set A (left)", color=S.colorA)[1]
    if c:
        S.colorA = c
        btn_colorA.configure(bg=c, activebackground=c)
        refresh_sheet_ui()

def pick_colorB():
    c = colorchooser.askcolor(title="Pick color for Set B (right)", color=S.colorB)[1]
    if c:
        S.colorB = c
        btn_colorB.configure(bg=c, activebackground=c)
        refresh_sheet_ui()

def update_alpha_labely():
    try:
        a = float(entry_alpha.get())
        ly = float(entry_labely.get())
        if not (0.0 <= a <= 1.0): raise ValueError
        if ly < 0.9 or ly > 1.6:   raise ValueError
        S.alpha = a
        S.label_y = ly
        refresh_sheet_ui()
    except Exception:
        messagebox.showerror("Invalid values", "Alpha must be 0–1 and label height ~0.9–1.6")

def prev_sheet():
    if not S.has_data(): return
    S.idx = (S.idx - 1) % len(S.sheets)
    refresh_sheet_ui()

def next_sheet():
    if not S.has_data(): return
    S.idx = (S.idx + 1) % len(S.sheets)
    refresh_sheet_ui()

def save_current_sheet():
    if not S.has_data(): return
    name = S.sheets[S.idx]
    df = S.dfs[name]
    la, lb = S.labels[name]
    setA, setB, uA, uB, sh = compute_sets(df.iloc[:,0], df.iloc[:,1])

    ts = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    base_dir = os.path.join(os.path.dirname(S.wb_path), "Venn_Outputs")
    os.makedirs(base_dir, exist_ok=True)

    la_safe = sanitize_filename(la); lb_safe = sanitize_filename(lb)
    base = f"{sanitize_filename(name)}__{la_safe}_vs_{lb_safe}_{ts}"

    # save PNG
    fig2, ax2 = plt.subplots(figsize=(6,6))
    draw_symmetric(ax2, setA, setB, la, lb, S.colorA, S.colorB, S.alpha, S.label_y)
    png_path = os.path.join(base_dir, base + ".png")
    fig2.savefig(png_path, dpi=150, bbox_inches="tight")
    plt.close(fig2)

    # save CSV
    csv_path = os.path.join(base_dir, base + ".csv")
    to_results_df(la, lb, uA, uB, sh).to_csv(csv_path, index=False)

    messagebox.showinfo("Saved", f"Saved PNG:\n{png_path}\n\nSaved CSV:\n{csv_path}")

def save_all_sheets():
    if not S.has_data(): return
    ts = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    base_dir = os.path.join(os.path.dirname(S.wb_path), "Venn_Outputs")
    os.makedirs(base_dir, exist_ok=True)

    # optional combined workbook
    combined = combined_var.get() == 1
    writer = None
    xlsx_path = None
    if combined:
        try:
            xlsx_path = os.path.join(base_dir, f"venn_batch_results_{ts}.xlsx")
            writer = pd.ExcelWriter(xlsx_path, engine="openpyxl")
        except Exception:
            writer = None
            messagebox.showwarning("Combined Excel",
                                   "openpyxl not available; skipping combined workbook.\nRun: pip install openpyxl")

    done = 0
    for name in S.sheets:
        df = S.dfs[name]
        if df.shape[1] < 2: continue
        la, lb = S.labels[name]
        setA, setB, uA, uB, sh = compute_sets(df.iloc[:,0], df.iloc[:,1])

        la_safe = sanitize_filename(la); lb_safe = sanitize_filename(lb)
        base = f"{sanitize_filename(name)}__{la_safe}_vs_{lb_safe}_{ts}"

        # PNG
        fig2, ax2 = plt.subplots(figsize=(6,6))
        draw_symmetric(ax2, setA, setB, la, lb, S.colorA, S.colorB, S.alpha, S.label_y)
        png_path = os.path.join(base_dir, base + ".png")
        fig2.savefig(png_path, dpi=150, bbox_inches="tight")
        plt.close(fig2)

        # CSV
        csv_path = os.path.join(base_dir, base + ".csv")
        df_out = to_results_df(la, lb, uA, uB, sh)
        df_out.to_csv(csv_path, index=False)

        if writer is not None:
            safe_sheet = sanitize_filename(name)[:31] or f"Sheet{done+1}"
            try:
                df_out.to_excel(writer, sheet_name=safe_sheet, index=False)
            except Exception:
                pass
        done += 1

    if writer is not None:
        try:
            writer.close()
            messagebox.showinfo("Saved",
                                f"Saved {done} sheet(s).\nCombined workbook:\n{xlsx_path}")
        except Exception:
            messagebox.showwarning("Combined Excel", "Failed to finalize combined workbook.")
    else:
        messagebox.showinfo("Saved", f"Saved {done} sheet(s) PNG+CSV in:\n{base_dir}")

# ---------- build UI ----------
root = tk.Tk()
root.title("Gene Venn (Excel → Multi-sheet, Symmetric)")

# File + sheet header
frm_top = tk.Frame(root); frm_top.pack(padx=8, pady=6, fill="x")
tk.Button(frm_top, text="Load Excel Workbook", command=load_workbook).grid(row=0, column=0, padx=6, sticky="w")
lbl_file = tk.Label(frm_top, text="Workbook: —"); lbl_file.grid(row=0, column=1, padx=6, sticky="w")
lbl_sheet = tk.Label(frm_top, text="Sheet: —"); lbl_sheet.grid(row=1, column=0, columnspan=2, padx=6, sticky="w")
lbl_preview = tk.Label(frm_top, text="Preview — Unique A: - | Shared: - | Unique B: -")
lbl_preview.grid(row=2, column=0, columnspan=3, padx=6, sticky="w")

# Labels / colors / alpha / label height
frm_ctrl = tk.Frame(root); frm_ctrl.pack(padx=8, pady=6, fill="x")
tk.Label(frm_ctrl, text="Label A").grid(row=0, column=0, sticky="e")
entry_labelA = tk.Entry(frm_ctrl, width=30); entry_labelA.grid(row=0, column=1, padx=4)
tk.Label(frm_ctrl, text="Label B").grid(row=1, column=0, sticky="e")
entry_labelB = tk.Entry(frm_ctrl, width=30); entry_labelB.grid(row=1, column=1, padx=4)
tk.Button(frm_ctrl, text="Apply Labels", command=apply_label_changes).grid(row=0, column=2, padx=6)
tk.Button(frm_ctrl, text="Reset to Excel Headers", command=reset_labels_to_headers).grid(row=1, column=2, padx=6)

tk.Label(frm_ctrl, text="Color A").grid(row=0, column=3, sticky="e")
btn_colorA = tk.Button(frm_ctrl, text="Pick…", width=8, command=pick_colorA, bg=S.colorA, activebackground=S.colorA)
btn_colorA.grid(row=0, column=4, padx=4)
tk.Label(frm_ctrl, text="Color B").grid(row=1, column=3, sticky="e")
btn_colorB = tk.Button(frm_ctrl, text="Pick…", width=8, command=pick_colorB, bg=S.colorB, activebackground=S.colorB)
btn_colorB.grid(row=1, column=4, padx=4)

tk.Label(frm_ctrl, text="Alpha (0–1)").grid(row=0, column=5, sticky="e")
entry_alpha = tk.Entry(frm_ctrl, width=6); entry_alpha.insert(0, str(S.alpha))
entry_alpha.grid(row=0, column=6, padx=4)
tk.Label(frm_ctrl, text="Label Height").grid(row=1, column=5, sticky="e")
entry_labely = tk.Entry(frm_ctrl, width=6); entry_labely.insert(0, str(S.label_y))
entry_labely.grid(row=1, column=6, padx=4)
tk.Button(frm_ctrl, text="Apply Alpha/Height", command=update_alpha_labely).grid(row=0, column=7, rowspan=2, padx=6)

# Navigation + save
frm_nav = tk.Frame(root); frm_nav.pack(padx=8, pady=6, fill="x")
tk.Button(frm_nav, text="⟨ Prev", width=10, command=prev_sheet).grid(row=0, column=0, padx=4)
tk.Button(frm_nav, text="Next ⟩", width=10, command=next_sheet).grid(row=0, column=1, padx=4)
tk.Button(frm_nav, text="Save Current (PNG+CSV)", command=save_current_sheet).grid(row=0, column=2, padx=10)
combined_var = tk.IntVar(value=1)
tk.Checkbutton(frm_nav, text="Also write combined Excel", variable=combined_var).grid(row=0, column=3, padx=8)
tk.Button(frm_nav, text="Save ALL Sheets", command=save_all_sheets).grid(row=0, column=4, padx=10)

# Plot canvas
frm_plot = tk.Frame(root); frm_plot.pack(padx=8, pady=6)
fig, ax = plt.subplots(figsize=(5.8,5.8))
canvas = FigureCanvasTkAgg(fig, master=frm_plot)
canvas.get_tk_widget().pack()

# Output panel
tk.Label(root, text="Preview lists (first 50 each):").pack(anchor="w", padx=8)
output_box = scrolledtext.ScrolledText(root, width=100, height=14, state="disabled")
output_box.pack(padx=8, pady=6)

root.mainloop()
