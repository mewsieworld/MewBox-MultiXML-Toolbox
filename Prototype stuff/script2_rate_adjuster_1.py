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

# ─── CSV parsers ─────────────────────────────────────────────────────────────
def parse_csv_line(line):
    result, cur, inq = [], "", False
    for c in line:
        if c == '"': inq = not inq; continue
        if c == ',' and not inq: result.append(cur); cur = ""; continue
        cur += c
    result.append(cur)
    return result

def parse_box_csv(text):
    """Returns {box_id_str: box_name_str} from either a 3-col-group CSV or a 2-col ID/BoxName CSV."""
    lines = [l for l in text.strip().splitlines() if l.strip()]
    if not lines:
        return {}
    headers = parse_csv_line(lines[0])
    box_map = {}

    # Detect 2-col simple CSV
    if len(headers) == 2 and headers[0].strip().upper() == "ID":
        for row in lines[1:]:
            cols = parse_csv_line(row)
            if len(cols) >= 2 and cols[0].strip().isdigit():
                box_map[cols[0].strip()] = cols[1].strip()
        return box_map

    # 3-col-group CSV (ID | Level/Rate | BoxName repeating)
    col_idx = 0
    while col_idx < len(headers):
        box_name = headers[col_idx + 2].strip() if col_idx + 2 < len(headers) else ""
        if box_name:
            # First row's ID for this group IS the parent box ID
            for row in lines[1:]:
                cols = parse_csv_line(row)
                pid = cols[col_idx].strip() if col_idx < len(cols) else ""
                if pid and pid.isdigit():
                    # We only want the box itself, not contents here —
                    # just record the mapping of box_name → IDs listed in rows
                    # For Script 2 we actually want the box IDs that ARE the parent boxes.
                    # The parent boxes are identified by their presence as COLUMN HEADERS (box_name).
                    # We'll store first non-empty id per group as the box's own ID IF it matches
                    # the box name column (i.e. col[i+2] == box_name header).
                    inner_name = cols[col_idx + 2].strip() if col_idx + 2 < len(cols) else ""
                    if inner_name == box_name:
                        box_map[pid] = box_name
                    break
        col_idx += 3

    # Fallback: collect ALL unique ids from each group header column
    if not box_map:
        col_idx = 0
        while col_idx < len(headers):
            box_name = headers[col_idx + 2].strip() if col_idx + 2 < len(headers) else ""
            if box_name:
                for row in lines[1:]:
                    cols = parse_csv_line(row)
                    pid = cols[col_idx].strip() if col_idx < len(cols) else ""
                    if pid and pid.isdigit():
                        box_map[pid] = box_name
                        break
            col_idx += 3

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

def adjust_row_to_100pct(block):
    """Set Type=2, DropCnt=real item count, DropRate=100 and ItemCnt=1 for real items, 0 for empty."""
    real_count = count_real_drops(block)
    block = set_tag_val(block, "Type",    "2")
    block = set_tag_val(block, "DropCnt", str(real_count))
    for i in range(20):
        id_val   = get_tag_val(block, f"DropId_{i}")
        has_item = bool(id_val and id_val != "0")
        block = set_tag_val(block, f"DropRate_{i}", "100" if has_item else "0")
        block = set_tag_val(block, f"ItemCnt_{i}",  "1"   if has_item else "0")
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

        def process():
            if not self.csv_text:
                messagebox.showwarning("Missing", "Please load a CSV first.")
                return
            if not self.xml_text:
                messagebox.showwarning("Missing", "Please load the PresentItemParam2.xml first.")
                return
            box_map = parse_box_csv(self.csv_text)
            if not box_map:
                messagebox.showerror("Error", "No box IDs could be parsed from the CSV.")
                return
            self._process(box_map)

        tk.Button(self, text="▶  Process →", command=process,
                  bg="#a6e3a1", fg="#1e1e2e", font=("Consolas", 12, "bold"),
                  relief="flat", padx=20, pady=8).pack(pady=20)

    def _process(self, box_map):
        all_rows = ROW_PATTERN.findall(self.xml_text)
        modified_rows = []
        csv_rows      = []   # for Script 3 export
        matched_count = 0

        for row in all_rows:
            row_id = get_tag_val(row, "Id")
            if row_id in box_map:
                new_row = adjust_row_to_100pct(row)
                modified_rows.append(new_row)
                # Build CSV row for Script 3
                drop_ids = get_all_drop_ids(row)
                csv_rows.append([row_id, box_map[row_id]] + drop_ids)
                matched_count += 1
            else:
                modified_rows.append(row)   # keep unmodified

        if matched_count == 0:
            messagebox.showwarning("No Matches",
                "None of the box IDs from the CSV were found in the XML.\n"
                "Check that your CSV IDs match the <Id> values in the XML.")
            return

        self._build_output_screen(modified_rows, csv_rows, matched_count)

    def _build_output_screen(self, modified_rows, csv_rows, matched_count):
        self._clear()
        tk.Label(self, text=f"Done — {matched_count} box(es) adjusted to 100% / Distributive",
                 font=("Consolas", 12, "bold"), bg="#1e1e2e", fg="#a6e3a1").pack(pady=12)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=12, pady=4)

        xml_content = "\n\n".join(modified_rows)

        def make_tab(title, text_content, default_filename):
            frm = tk.Frame(nb, bg="#1e1e2e")
            nb.add(frm, text=title)
            txt = scrolledtext.ScrolledText(frm, font=("Consolas", 9),
                                            bg="#181825", fg="#cdd6f4")
            txt.pack(fill="both", expand=True, padx=4, pady=4)
            txt.insert("1.0", text_content)
            txt.config(state="disabled")

            def copy_all():
                self.clipboard_clear(); self.clipboard_append(text_content)
                messagebox.showinfo("Copied", "Copied to clipboard.")

            def save():
                path = filedialog.asksaveasfilename(initialfile=default_filename,
                                                    defaultextension=".xml",
                                                    filetypes=[("XML","*.xml"),("CSV","*.csv"),("All","*.*")])
                if path:
                    with open(path, "w", encoding="utf-8") as f:
                        f.write(text_content)
                    messagebox.showinfo("Saved", f"Saved to {path}")

            brow = tk.Frame(frm, bg="#1e1e2e")
            brow.pack(fill="x")
            tk.Button(brow, text="📋 Copy", command=copy_all,
                      bg="#313244", fg="#cdd6f4", font=("Consolas", 9),
                      relief="flat", padx=10, pady=4).pack(side="left", padx=6, pady=4)
            tk.Button(brow, text="💾 Save As…", command=save,
                      bg="#a6e3a1", fg="#1e1e2e", font=("Consolas", 9),
                      relief="flat", padx=10, pady=4).pack(side="left", padx=6, pady=4)

        # Build CSV content for Script 3
        csv_lines = ["BoxID,BoxName,ItemIDs..."]
        max_items = max((len(r) - 2 for r in csv_rows), default=0)
        # Header with numbered item columns
        header = ["BoxID", "BoxName"] + [f"Item{i+1}_ID" for i in range(max_items)]
        csv_lines = [",".join(header)]
        for row in csv_rows:
            csv_lines.append(",".join(str(x) for x in row))
        csv_content = "\n".join(csv_lines)

        make_tab("Modified PresentItemParam2 rows", xml_content, "PresentItemParam2_100pct.xml")
        make_tab("Box Contents CSV (for Script 3)", csv_content, "box_contents_for_script3.csv")

        tk.Button(self, text="◀  Start Over", command=self._build_load_screen,
                  bg="#313244", fg="#cdd6f4", font=("Consolas", 10),
                  relief="flat", padx=12, pady=6).pack(pady=8)

    def _clear(self):
        for w in self.winfo_children():
            w.destroy()


import os

if __name__ == "__main__":
    app = RateAdjusterApp()
    app.mainloop()
