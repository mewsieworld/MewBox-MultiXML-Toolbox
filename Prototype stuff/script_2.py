"""
RecycleExceptItem XML Generator
Generates <ROW><ItemID>...</ItemID><Comment><![CDATA[...]]></Comment></ROW> entries.
Run: python script.py
"""

import csv
import io
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

BG       = "#1e1e2e"
BG2      = "#181825"
BG3      = "#313244"
FG       = "#cdd6f4"
ACCENT   = "#cba6f7"
GREEN    = "#a6e3a1"
BLUE     = "#89b4fa"
MUTED    = "#6c7086"
FONT     = ("Consolas", 10)
FONT_SM  = ("Consolas", 9)
FONT_LG  = ("Consolas", 13, "bold")


def generate_xml(rows):
    lines = []
    for item_id, name in rows:
        lines.append(
            f"<ROW>\n"
            f"<ItemID>{item_id}</ItemID>\n"
            f"<Comment><![CDATA[{name}]]></Comment>\n"
            f"</ROW>"
        )
    return "\n".join(lines)


def parse_csv_text(text):
    rows = []
    reader = csv.DictReader(io.StringIO(text.strip()))
    if not reader.fieldnames:
        return None, "No columns found."
    id_col   = next((h for h in reader.fieldnames if h.strip().lower() == "id"), None)
    name_col = next((h for h in reader.fieldnames if h.strip().lower() in ("name", "comment")), None)
    if not id_col:
        return None, "CSV must have an 'ID' column."
    if not name_col:
        return None, "CSV must have a 'Name' or 'Comment' column."
    for row in reader:
        item_id = row[id_col].strip()
        name    = row[name_col].strip()
        if item_id and name:
            rows.append((item_id, name))
    if not rows:
        return None, "No valid rows found."
    return rows, None


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RecycleExceptItem XML Generator")
        self.geometry("900x700")
        self.configure(bg=BG)
        self.resizable(True, True)

        # Manual row data: list of (id_var, name_var)
        self.manual_rows = []

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TNotebook",     background=BG,  borderwidth=0)
        style.configure("TNotebook.Tab", background=BG3, foreground=MUTED,
                        padding=[14, 6], font=FONT)
        style.map("TNotebook.Tab",
                  background=[("selected", "#45475a")],
                  foreground=[("selected", ACCENT)])
        style.configure("TFrame", background=BG)
        style.configure("Vertical.TScrollbar", background=BG3, troughcolor=BG2,
                        arrowcolor=FG, bordercolor=BG)

        # ── Header ──────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=BG2)
        hdr.pack(fill="x")
        tk.Label(hdr, text="RecycleExceptItem  XML Generator",
                 font=FONT_LG, bg=BG2, fg=ACCENT, pady=10, padx=16).pack(side="left")

        # ── Notebook ─────────────────────────────────────────────────────────
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=False, padx=12, pady=(8, 0))

        tab_file   = ttk.Frame(nb)
        tab_manual = ttk.Frame(nb)
        nb.add(tab_file,   text="  Load CSV File  ")
        nb.add(tab_manual, text="  Manual Entry   ")

        self._build_file_tab(tab_file)
        self._build_manual_tab(tab_manual)

        # ── Output ───────────────────────────────────────────────────────────
        out_lf = tk.LabelFrame(self, text="  XML Output  ", bg=BG, fg=BLUE,
                               font=("Consolas", 10, "bold"), bd=1, relief="groove")
        out_lf.pack(fill="both", expand=True, padx=12, pady=(6, 0))

        self.output_text = scrolledtext.ScrolledText(
            out_lf, font=FONT_SM, bg=BG2, fg="#a0c8e0",
            insertbackground=ACCENT, relief="flat", wrap="none", height=12)
        self.output_text.pack(fill="both", expand=True, padx=4, pady=4)

        # ── Bottom bar ───────────────────────────────────────────────────────
        bot = tk.Frame(self, bg=BG)
        bot.pack(fill="x", padx=12, pady=8)

        self.status_lbl = tk.Label(bot, text="", bg=BG, fg=MUTED, font=FONT_SM)
        self.status_lbl.pack(side="left")

        tk.Button(bot, text="💾  Save  RecycleExceptItem.xml",
                  command=self._save,
                  bg="#2a2a10", fg=ACCENT, font=("Consolas", 10, "bold"),
                  relief="flat", padx=14, pady=6, cursor="hand2",
                  activebackground="#3a3a18", activeforeground=ACCENT).pack(side="right")

        tk.Button(bot, text="📋  Copy XML",
                  command=self._copy,
                  bg=BG3, fg=FG, font=FONT_SM,
                  relief="flat", padx=12, pady=6, cursor="hand2",
                  activebackground="#45475a").pack(side="right", padx=(0, 8))

    # ══════════════════════════════════════════════════════════════════════
    # FILE TAB
    # ══════════════════════════════════════════════════════════════════════
    def _build_file_tab(self, parent):
        inner = tk.Frame(parent, bg=BG)
        inner.pack(expand=True, pady=18)

        tk.Label(inner,
                 text="Load a CSV file with columns:  ID  |  Name  (or Comment)",
                 bg=BG, fg=FG, font=FONT).pack(pady=(0, 12))

        btn_row = tk.Frame(inner, bg=BG)
        btn_row.pack()

        self.file_label = tk.Label(inner, text="", bg=BG, fg=GREEN, font=FONT_SM)
        self.file_label.pack(pady=(10, 0))

        tk.Button(btn_row, text="📂  Browse CSV File",
                  command=self._browse_file,
                  bg=BG3, fg=FG, font=FONT,
                  relief="flat", padx=14, pady=8, cursor="hand2",
                  activebackground="#45475a").pack(side="left", padx=8)

        tk.Button(btn_row, text="📋  Paste CSV Text",
                  command=self._paste_csv,
                  bg=BG3, fg=FG, font=FONT,
                  relief="flat", padx=14, pady=8, cursor="hand2",
                  activebackground="#45475a").pack(side="left", padx=8)

    def _browse_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, encoding="utf-8-sig") as f:
                text = f.read()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        rows, err = parse_csv_text(text)
        if err:
            messagebox.showerror("Parse Error", err)
            return
        self.file_label.config(text=f"✓  {os.path.basename(path)}  ({len(rows)} rows)")
        self._render_xml(rows)

    def _paste_csv(self):
        win = tk.Toplevel(self)
        win.title("Paste CSV")
        win.geometry("600x380")
        win.configure(bg=BG)
        tk.Label(win, text="Paste CSV content below:", bg=BG, fg=FG, font=FONT).pack(
            anchor="w", padx=12, pady=(10, 4))
        txt = scrolledtext.ScrolledText(win, font=FONT_SM, bg=BG2, fg="#c8e0a0",
                                         insertbackground=ACCENT)
        txt.pack(fill="both", expand=True, padx=12, pady=4)
        txt.insert("1.0", "ID,Name\n")

        def confirm():
            rows, err = parse_csv_text(txt.get("1.0", "end"))
            if err:
                messagebox.showerror("Parse Error", err)
                return
            self.file_label.config(text=f"✓  Pasted CSV  ({len(rows)} rows)")
            self._render_xml(rows)
            win.destroy()

        tk.Button(win, text="Generate XML", command=confirm,
                  bg=GREEN, fg=BG2, font=("Consolas", 10, "bold"),
                  relief="flat", padx=14, pady=6).pack(pady=8)

    # ══════════════════════════════════════════════════════════════════════
    # MANUAL TAB
    # ══════════════════════════════════════════════════════════════════════
    def _build_manual_tab(self, parent):
        # Header row labels
        hdr = tk.Frame(parent, bg=BG2)
        hdr.pack(fill="x", padx=0, pady=(0, 1))
        tk.Label(hdr, text="  #",    width=4,  anchor="w", bg=BG2, fg=MUTED, font=FONT_SM).pack(side="left", padx=(8,0))
        tk.Label(hdr, text="Item ID",          width=14, anchor="w", bg=BG2, fg=BLUE,  font=("Consolas", 9, "bold")).pack(side="left", padx=4)
        tk.Label(hdr, text="Name / Comment",               anchor="w", bg=BG2, fg=BLUE,  font=("Consolas", 9, "bold")).pack(side="left", padx=4)

        # Scrollable rows area
        canvas_frame = tk.Frame(parent, bg=BG)
        canvas_frame.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(canvas_frame, bg=BG, highlightthickness=0, height=160)
        vsb = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.rows_frame = tk.Frame(self.canvas, bg=BG)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.rows_frame, anchor="nw")

        self.rows_frame.bind("<Configure>", lambda e: self.canvas.configure(
            scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(
            self.canvas_window, width=e.width))
        self.canvas.bind_all("<MouseWheel>",
            lambda e: self.canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        # Bottom controls
        ctrl = tk.Frame(parent, bg=BG2)
        ctrl.pack(fill="x", pady=(2, 0))

        tk.Button(ctrl, text="＋  Add Row",
                  command=self._add_manual_row,
                  bg=BG3, fg=FG, font=FONT_SM,
                  relief="flat", padx=10, pady=5, cursor="hand2",
                  activebackground="#45475a").pack(side="left", padx=8, pady=6)

        tk.Button(ctrl, text="▶  Generate XML",
                  command=self._generate_from_manual,
                  bg=GREEN, fg=BG2, font=("Consolas", 10, "bold"),
                  relief="flat", padx=14, pady=5, cursor="hand2",
                  activebackground="#b5f0b0").pack(side="right", padx=8, pady=6)

        tk.Button(ctrl, text="🗑  Clear All",
                  command=self._clear_manual_rows,
                  bg=BG3, fg=MUTED, font=FONT_SM,
                  relief="flat", padx=10, pady=5, cursor="hand2").pack(side="right", padx=(0, 4), pady=6)

        # Seed with a few blank rows
        for _ in range(5):
            self._add_manual_row()

    def _add_manual_row(self, id_val="", name_val=""):
        row_num = len(self.manual_rows) + 1
        id_var   = tk.StringVar(value=id_val)
        name_var = tk.StringVar(value=name_val)
        self.manual_rows.append((id_var, name_var))

        row_frame = tk.Frame(self.rows_frame, bg=BG if row_num % 2 == 0 else BG2)
        row_frame.pack(fill="x")

        tk.Label(row_frame, text=f"  {row_num}", width=4, anchor="w",
                 bg=row_frame["bg"], fg=MUTED, font=FONT_SM).pack(side="left", padx=(8, 0), pady=2)

        tk.Entry(row_frame, textvariable=id_var, width=14,
                 bg=BG3, fg="#f9e2af", insertbackground=ACCENT,
                 font=FONT_SM, relief="flat").pack(side="left", padx=4, pady=3)

        tk.Entry(row_frame, textvariable=name_var,
                 bg=BG3, fg="#cdd6f4", insertbackground=ACCENT,
                 font=FONT_SM, relief="flat").pack(side="left", fill="x", expand=True, padx=4, pady=3)

        tk.Button(row_frame, text="✕",
                  command=lambda f=row_frame, r=(id_var, name_var): self._delete_row(f, r),
                  bg=row_frame["bg"], fg="#f38ba8", font=("Consolas", 9),
                  relief="flat", padx=4, cursor="hand2",
                  activebackground=BG3).pack(side="right", padx=4)

    def _delete_row(self, frame, row_tuple):
        if row_tuple in self.manual_rows:
            self.manual_rows.remove(row_tuple)
        frame.destroy()
        self._renumber_rows()

    def _renumber_rows(self):
        for i, child in enumerate(self.rows_frame.winfo_children()):
            bg = BG if (i + 1) % 2 == 0 else BG2
            child.configure(bg=bg)
            labels = [w for w in child.winfo_children() if isinstance(w, tk.Label)]
            if labels:
                labels[0].configure(text=f"  {i+1}", bg=bg)
            for w in child.winfo_children():
                if not isinstance(w, tk.Entry):
                    w.configure(bg=bg)

    def _clear_manual_rows(self):
        for w in self.rows_frame.winfo_children():
            w.destroy()
        self.manual_rows.clear()
        for _ in range(5):
            self._add_manual_row()

    def _generate_from_manual(self):
        rows = []
        for id_var, name_var in self.manual_rows:
            item_id = id_var.get().strip()
            name    = name_var.get().strip()
            if item_id and name:
                rows.append((item_id, name))
        if not rows:
            messagebox.showwarning("No Data", "Enter at least one ID and Name row.")
            return
        self._render_xml(rows)

    # ══════════════════════════════════════════════════════════════════════
    # Shared output helpers
    # ══════════════════════════════════════════════════════════════════════
    def _render_xml(self, rows):
        xml = generate_xml(rows)
        self.output_text.config(state="normal")
        self.output_text.delete("1.0", "end")
        self.output_text.insert("1.0", xml)
        self.status_lbl.config(text=f"{len(rows)} rows ready", fg=GREEN)

    def _copy(self):
        xml = self.output_text.get("1.0", "end").strip()
        if not xml:
            return
        self.clipboard_clear()
        self.clipboard_append(xml)
        self.status_lbl.config(text="Copied to clipboard ✓", fg=ACCENT)

    def _save(self):
        xml = self.output_text.get("1.0", "end").strip()
        if not xml:
            messagebox.showwarning("Nothing to save", "Generate XML first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xml",
            initialfile="RecycleExceptItem.xml",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")])
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            f.write(xml)
        self.status_lbl.config(text=f"Saved → {os.path.abspath(path)}", fg=GREEN)
        messagebox.showinfo("Saved", f"File saved to:\n{path}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
