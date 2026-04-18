"""
SCRIPT 2 – Box Rate Adjuster (100% / Distributive Converter)
─────────────────────────────────────────────────────────────
Workflow:
  1. Load a box-list CSV  (columns: ID | anything | ParentBoxName  — repeating groups of 3)
     OR a simple 2-col CSV (ID, BoxName) exported by Script 1.
  2. Load the existing PresentItemParam2.xml that contains <ROW> entries for those boxes.
  3. For every matching box:
       • DropCnt  → set to the number of items actually present in that row (non-zero IDs)
       • DropRate → set to 100 for every real item slot
       • Type     → set to 2 (Distributive)
  4. Export:
       • Modified PresentItemParam2.xml rows  (no <?xml header, just the <ROW> blocks)
       • A CSV of every box's contents  (BoxID, BoxName, Item1_ID, Item2_ID, …)
         compatible with Script 3 (optional Ticket/Cost column can be added manually)

Requirements: Python 3.x  (standard library only)
Run: python script2_rate_adjuster.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import csv, io, re, xml.etree.ElementTree as ET


# ─── Flexible CSV parser (works for any group width) ─────────────────────────
# Header groups can be:  ID, ParentBoxName
#                        ID, Level, ParentBoxName
#                        ID, Rate, ParentBoxName
#                        ID, Level, Rate, ParentBoxName   (or any order of extras)
# The FIRST column of each group is always the item ID.
# The LAST column of each group is always the parent box name (non-numeric header).
# Any middle columns (Level, Rate, Lv, Luck, etc.) are collected but not required.

import csv as _csv, io as _io

_SKIP_HEADERS = {"id", "level", "rate", "lv", "luck", "lvl", "chance", "prob"}

def _is_box_name_header(h):
    """True if this header looks like a parent-box name rather than a field label."""
    h = h.strip()
    return bool(h) and h.lower() not in _SKIP_HEADERS and not h.isdigit()

def parse_grouped_csv(text):
    """
    Parse a CSV with repeating column groups.
    Returns a list of:
      { "box_name": str,
        "items": [{"id": str, "name": str, "extra": [str, ...]}, ...] }
    """
    reader = _csv.reader(_io.StringIO(text))
    rows   = list(reader)
    if not rows:
        return []

    headers = [h.strip() for h in rows[0]]

    # Find every column index whose header is a parent-box name
    name_col_indices = [i for i, h in enumerate(headers) if _is_box_name_header(h)]
    if not name_col_indices:
        return []

    # Build group metadata
    groups_meta = []
    prev_end = -1
    for nc in name_col_indices:
        id_col     = prev_end + 1       # first col of this group = ID
        middle     = list(range(id_col + 1, nc))  # optional Level/Rate cols
        groups_meta.append({
            "box_name":    headers[nc],
            "id_col":      id_col,
            "name_col":    nc,
            "middle_cols": middle,
        })
        prev_end = nc

    # Parse data rows
    results = []
    for gm in groups_meta:
        items = []
        for row in rows[1:]:
            id_val = row[gm["id_col"]].strip() if gm["id_col"] < len(row) else ""
            if not id_val or not id_val.isdigit():
                continue
            nm_val = row[gm["name_col"]].strip() if gm["name_col"] < len(row) else ""
            extras = [row[c].strip() for c in gm["middle_cols"] if c < len(row)]
            # Try to extract a numeric rate from extras (last numeric extra = rate)
            rate = None
            for ex in reversed(extras):
                clean = ex.replace("%","").strip()
                if clean.isdigit():
                    rate = int(clean)
                    break
            items.append({"id": id_val, "name": nm_val, "extra": extras, "rate": rate})
        if items:
            results.append({"box_name": gm["box_name"], "items": items})
    return results

def parse_box_csv_for_script2(text):
    """
    Returns {mini_box_id_str: mini_box_name_str} for ALL item rows across ALL groups.
    The header row identifies the parent box (ignored for matching purposes).
    Every numeric ID in every data row is a mini-box to be modified.
    Duplicate IDs across groups are fine — last name wins (they're the same box).
    """
    groups = parse_grouped_csv(text)
    box_map = {}
    for g in groups:
        for item in g["items"]:
            box_map[item["id"]] = item["name"]
    return box_map

# ─── XML helpers ─────────────────────────────────────────────────────────────
ROW_PATTERN     = re.compile(r'<ROW>.*?</ROW>', re.DOTALL)
CDATA_PATTERN   = re.compile(r'<!\[CDATA\[(.*?)\]\]>', re.DOTALL)

def get_tag_val(block, tag):
    m = re.search(rf'<{tag}>(.*?)</{tag}>', block, re.DOTALL)
    if m:
        inner = m.group(1)
        cd = CDATA_PATTERN.search(inner)
        return cd.group(1) if cd else inner.strip()
    return ""

def set_tag_val(block, tag, new_val):
    return re.sub(rf'<{tag}>.*?</{tag}>', f'<{tag}>{new_val}</{tag}>', block, flags=re.DOTALL)

def count_real_drops(block):
    """Count DropId_# entries that are non-zero."""
    ids = re.findall(r'<DropId_\d+>(\d+)</DropId_\d+>', block)
    return sum(1 for x in ids if x != "0")

def get_all_drop_ids(block):
    """Return list of non-zero DropId values in order."""
    pairs = re.findall(r'<DropId_(\d+)>(\d+)</DropId_\d+>', block)
    result = []
    for _, val in sorted(pairs, key=lambda x: int(x[0])):
        if val != "0":
            result.append(val)
    return result

def adjust_row(block, rate, count):
    """
    For every slot that has a non-zero DropId_#:
      - Set DropRate_# to `rate`
      - Set ItemCnt_#  to `count`
    Everything else (Type, DropCnt, empty slots) is left exactly as-is.
    """
    for i in range(20):
        id_val = get_tag_val(block, f"DropId_{i}")
        if id_val and id_val != "0":
            block = set_tag_val(block, f"DropRate_{i}", str(rate))
            block = set_tag_val(block, f"ItemCnt_{i}",  str(count))
    return block

# ─── App ─────────────────────────────────────────────────────────────────────
class RateAdjusterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Script 2 – Box Rate Adjuster (→ 100% / Distributive)")
        self.geometry("860x620")
        self.configure(bg="#1e1e2e")
        self.csv_text  = ""
        self.xml_text  = ""
        self._build_load_screen()

    def _build_load_screen(self):
        self._clear()
        tk.Label(self, text="BOX RATE ADJUSTER", font=("Consolas", 18, "bold"),
                 bg="#1e1e2e", fg="#89dceb").pack(pady=(30, 5))
        tk.Label(self,
                 text="Step 1 — Load your box CSV\n(3-col groups: ID | Level/Rate | BoxName,  OR  2-col: ID | BoxName)",
                 bg="#1e1e2e", fg="#a6adc8", font=("Consolas", 10), justify="center").pack(pady=8)

        # CSV section
        csv_frm = tk.LabelFrame(self, text="  Box CSV  ", bg="#1e1e2e", fg="#89b4fa",
                                font=("Consolas", 10, "bold"), bd=1, relief="groove")
        csv_frm.pack(fill="x", padx=30, pady=6)

        csv_status = tk.StringVar(value="No file loaded")
        tk.Label(csv_frm, textvariable=csv_status, bg="#1e1e2e",
                 fg="#6c7086", font=("Consolas", 9)).pack(side="left", padx=10)

        def load_csv():
            path = filedialog.askopenfilename(filetypes=[("CSV","*.csv"),("All","*.*")])
            if not path: return
            with open(path, encoding="utf-8-sig") as f:
                self.csv_text = f.read()
            csv_status.set(f"✓  {os.path.basename(path)}")
        tk.Button(csv_frm, text="📂 Load CSV", command=load_csv,
                  bg="#313244", fg="#cdd6f4", font=("Consolas", 9),
                  relief="flat", padx=10, pady=4).pack(side="right", padx=8, pady=6)

        import os
        # XML section
        tk.Label(self,
                 text="Step 2 — Load your existing PresentItemParam2.xml",
                 bg="#1e1e2e", fg="#a6adc8", font=("Consolas", 10)).pack(pady=(10, 2))

        xml_frm = tk.LabelFrame(self, text="  PresentItemParam2.xml  ", bg="#1e1e2e", fg="#89b4fa",
                                font=("Consolas", 10, "bold"), bd=1, relief="groove")
        xml_frm.pack(fill="x", padx=30, pady=6)

        xml_status = tk.StringVar(value="No file loaded")
        tk.Label(xml_frm, textvariable=xml_status, bg="#1e1e2e",
                 fg="#6c7086", font=("Consolas", 9)).pack(side="left", padx=10)

        def load_xml():
            path = filedialog.askopenfilename(filetypes=[("XML","*.xml"),("All","*.*")])
            if not path: return
            with open(path, encoding="utf-8-sig") as f:
                self.xml_text = f.read()
            xml_status.set(f"✓  {os.path.basename(path)}")
        tk.Button(xml_frm, text="📂 Load XML", command=load_xml,
                  bg="#313244", fg="#cdd6f4", font=("Consolas", 9),
                  relief="flat", padx=10, pady=4).pack(side="right", padx=8, pady=6)

        # ── Adjust Rate / Adjust Count ──────────────────────────────────────────
        adj_frm = tk.LabelFrame(self, text="  Adjustment Values  ", bg="#1e1e2e", fg="#89b4fa",
                                font=("Consolas", 10, "bold"), bd=1, relief="groove")
        adj_frm.pack(fill="x", padx=30, pady=6)

        def _make_num_row(parent, label, default):
            r = tk.Frame(parent, bg="#1e1e2e")
            r.pack(fill="x", padx=10, pady=4)
            tk.Label(r, text=label, width=18, anchor="w", bg="#1e1e2e",
                     fg="#cdd6f4", font=("Consolas", 9)).pack(side="left")
            var = tk.StringVar(value=str(default))
            ent = tk.Entry(r, textvariable=var, width=8,
                           bg="#313244", fg="#cdd6f4", insertbackground="#cdd6f4",
                           font=("Consolas", 9), relief="flat")
            ent.pack(side="left", padx=6)
            tk.Label(r, text="(1 – 32766)", bg="#1e1e2e", fg="#6c7086",
                     font=("Consolas", 8)).pack(side="left")
            return var

        rate_var  = _make_num_row(adj_frm, "Adjust Rate:",  100)
        count_var = _make_num_row(adj_frm, "Adjust Count:", 1)

        def process():
            if not self.csv_text:
                messagebox.showwarning("Missing", "Please load a CSV first.")
                return
            if not self.xml_text:
                messagebox.showwarning("Missing", "Please load the PresentItemParam2.xml first.")
                return
            box_map = parse_box_csv_for_script2(self.csv_text)
            if not box_map:
                messagebox.showerror("Error", "No box IDs could be parsed from the CSV.")
                return
            try:
                rate  = int(rate_var.get())
                count = int(count_var.get())
                if not (1 <= rate  <= 32766): raise ValueError
                if not (1 <= count <= 32766): raise ValueError
            except ValueError:
                messagebox.showerror("Invalid", "Rate and Count must both be integers between 1 and 32766.")
                return
            self._process(box_map, rate, count)

        tk.Button(self, text="▶  Process →", command=process,
                  bg="#a6e3a1", fg="#1e1e2e", font=("Consolas", 12, "bold"),
                  relief="flat", padx=20, pady=8).pack(pady=20)

    def _process(self, box_map, rate, count):
        csv_rows      = []
        matched_count = 0

        full_xml = self.xml_text

        def replace_row(m):
            nonlocal matched_count
            row    = m.group(0)
            row_id = get_tag_val(row, "Id")
            if row_id in box_map:
                new_row  = adjust_row(row, rate, count)
                drop_ids = get_all_drop_ids(new_row)
                csv_rows.append([row_id, box_map[row_id]] + drop_ids)
                matched_count += 1
                return new_row
            return row  # completely untouched

        full_xml_modified = ROW_PATTERN.sub(replace_row, full_xml)

        if matched_count == 0:
            messagebox.showwarning("No Matches",
                "None of the box IDs from the CSV were found in the XML.\n"
                "Check that your CSV IDs match the <Id> values in the XML.")
            return

        self._build_output_screen(full_xml_modified, csv_rows, matched_count)

    def _build_output_screen(self, full_xml_modified, csv_rows, matched_count):
        self._clear()
        tk.Label(self, text=f"Done — {matched_count} box(es) adjusted",
                 font=("Consolas", 12, "bold"), bg="#1e1e2e", fg="#a6e3a1").pack(pady=12)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=12, pady=4)

        # CSV for Script 3
        header    = ["BoxID", "BoxName"] + [f"Item{i+1}_ID" for i in range(max((len(r)-2 for r in csv_rows), default=0))]
        csv_lines = [",".join(header)] + [",".join(str(x) for x in r) for r in csv_rows]
        csv_content = "\n".join(csv_lines)

        def make_tab(title, content, fname):
            frm = tk.Frame(nb, bg="#1e1e2e")
            nb.add(frm, text=title)
            # buttons anchored to bottom so always visible — closure-safe via default args
            br = tk.Frame(frm, bg="#1e1e2e")
            br.pack(side="bottom", fill="x")
            tk.Button(br, text="📋 Copy All",
                      command=lambda c=content: (self.clipboard_clear(),
                                                  self.clipboard_append(c),
                                                  messagebox.showinfo("Copied","Copied to clipboard.")),
                      bg="#313244", fg="#cdd6f4", font=("Consolas", 9),
                      relief="flat", padx=10, pady=4).pack(side="left", padx=6, pady=4)
            tk.Button(br, text="💾 Save As…",
                      command=lambda c=content, f=fname: self._save(c, f),
                      bg="#a6e3a1", fg="#1e1e2e", font=("Consolas", 9),
                      relief="flat", padx=10, pady=4).pack(side="left", padx=6, pady=4)
            txt = scrolledtext.ScrolledText(frm, font=("Consolas", 9),
                                            bg="#181825", fg="#cdd6f4")
            txt.pack(fill="both", expand=True, padx=4, pady=4)
            txt.insert("1.0", content)
            txt.config(state="disabled")

        make_tab("Full PresentItemParam2.xml (modified)", full_xml_modified, "PresentItemParam2_modified.xml")
        make_tab("Box Contents CSV (for Script 3)",       csv_content,       "box_contents_for_script3.csv")

        tk.Button(self, text="◀  Start Over", command=self._build_load_screen,
                  bg="#313244", fg="#cdd6f4", font=("Consolas", 10),
                  relief="flat", padx=12, pady=6).pack(pady=8)

    def _save(self, content, fname):
        path = filedialog.asksaveasfilename(initialfile=fname,
                   defaultextension=".xml",
                   filetypes=[("XML","*.xml"),("CSV","*.csv"),("All","*.*")])
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(content)
            messagebox.showinfo("Saved", f"Saved to {path}")

    def _clear(self):
        for w in self.winfo_children():
            w.destroy()


import os

if __name__ == "__main__":
    app = RateAdjusterApp()
    app.mainloop()
