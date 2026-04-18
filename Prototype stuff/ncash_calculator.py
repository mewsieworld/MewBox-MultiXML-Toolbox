"""
NCash ↔ Ticket Calculator
Formula: NCash = round(Tickets × 133)
         Tickets = NCash / 133
Run: python ncash_calculator.py
"""

import tkinter as tk
from tkinter import font as tkfont

class NCashCalc(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("NCash ↔ Ticket Calculator")
        self.configure(bg="#1e1e2e")
        self.resizable(False, False)

        pad = dict(bg="#1e1e2e")

        tk.Label(self, text="NCash  ↔  Ticket  Calculator",
                 font=("Consolas", 16, "bold"),
                 bg="#1e1e2e", fg="#cba6f7").pack(pady=(24, 4))
        tk.Label(self, text="Formula:  NCash = round( Tickets × 133 )",
                 font=("Consolas", 9), bg="#1e1e2e", fg="#6c7086").pack(pady=(0, 18))

        # ── Tickets → NCash ──────────────────────────────────────────────
        box_a = tk.LabelFrame(self, text="  Tickets  →  NCash  ",
                              bg="#1e1e2e", fg="#89b4fa",
                              font=("Consolas", 10, "bold"),
                              bd=1, relief="groove")
        box_a.pack(fill="x", padx=28, pady=6)

        row_a = tk.Frame(box_a, **pad); row_a.pack(padx=14, pady=10)
        tk.Label(row_a, text="Tickets:", width=10, anchor="w",
                 font=("Consolas", 11), bg="#1e1e2e", fg="#cdd6f4").pack(side="left")
        self.v_tickets = tk.StringVar()
        tk.Entry(row_a, textvariable=self.v_tickets, width=14,
                 bg="#313244", fg="#cdd6f4", insertbackground="#cdd6f4",
                 font=("Consolas", 13), relief="flat").pack(side="left", padx=8)
        tk.Label(row_a, text="=", font=("Consolas", 13),
                 bg="#1e1e2e", fg="#6c7086").pack(side="left", padx=4)
        self.lbl_ncash = tk.Label(row_a, text="—", width=14, anchor="w",
                                   font=("Consolas", 13, "bold"),
                                   bg="#1e1e2e", fg="#a6e3a1")
        self.lbl_ncash.pack(side="left", padx=4)
        tk.Label(row_a, text="NCash", font=("Consolas", 10),
                 bg="#1e1e2e", fg="#6c7086").pack(side="left")

        self.v_tickets.trace_add("write", self._calc_ncash)

        # ── NCash → Tickets ──────────────────────────────────────────────
        box_b = tk.LabelFrame(self, text="  NCash  →  Tickets  ",
                              bg="#1e1e2e", fg="#89b4fa",
                              font=("Consolas", 10, "bold"),
                              bd=1, relief="groove")
        box_b.pack(fill="x", padx=28, pady=6)

        row_b = tk.Frame(box_b, **pad); row_b.pack(padx=14, pady=10)
        tk.Label(row_b, text="NCash:", width=10, anchor="w",
                 font=("Consolas", 11), bg="#1e1e2e", fg="#cdd6f4").pack(side="left")
        self.v_ncash = tk.StringVar()
        tk.Entry(row_b, textvariable=self.v_ncash, width=14,
                 bg="#313244", fg="#cdd6f4", insertbackground="#cdd6f4",
                 font=("Consolas", 13), relief="flat").pack(side="left", padx=8)
        tk.Label(row_b, text="=", font=("Consolas", 13),
                 bg="#1e1e2e", fg="#6c7086").pack(side="left", padx=4)
        self.lbl_tickets = tk.Label(row_b, text="—", width=14, anchor="w",
                                     font=("Consolas", 13, "bold"),
                                     bg="#1e1e2e", fg="#f9e2af")
        self.lbl_tickets.pack(side="left", padx=4)
        tk.Label(row_b, text="Tickets", font=("Consolas", 10),
                 bg="#1e1e2e", fg="#6c7086").pack(side="left")

        self.v_ncash.trace_add("write", self._calc_tickets)

        tk.Label(self, text="", bg="#1e1e2e").pack(pady=8)
        self.geometry("480x310")

    def _calc_ncash(self, *_):
        try:
            t = float(self.v_tickets.get())
            self.lbl_ncash.config(text=f"{round(t * 133):,}")
        except:
            self.lbl_ncash.config(text="—")

    def _calc_tickets(self, *_):
        try:
            n = float(self.v_ncash.get())
            raw = n / 133
            # Show both rounded and exact if not whole
            rounded = round(raw, 4)
            if rounded == int(rounded):
                self.lbl_tickets.config(text=f"{int(rounded):,}")
            else:
                self.lbl_tickets.config(text=f"{rounded:,.4f}")
        except:
            self.lbl_tickets.config(text="—")

if __name__ == "__main__":
    NCashCalc().mainloop()
