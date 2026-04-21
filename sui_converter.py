"""
Michigan MiUI SUI Wage Report Converter
Converts QuickBooks Desktop Enterprise .xlsm SUI export to MiUI Delimited CSV format.

Requirements:
    pip install pandas openpyxl

To build .exe (run in WSL or Windows with Python installed):
    pip install pyinstaller
    pyinstaller --onefile --windowed --name "SUI_Converter" sui_converter.py

Usage: Run the script or .exe — a GUI window will appear.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import sys
import datetime
import re

try:
    import pandas as pd
except ImportError:
    tk.Tk().withdraw()
    messagebox.showerror("Missing Dependency", "pandas is required.\n\nRun: pip install pandas openpyxl")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Quarter helpers
# ---------------------------------------------------------------------------

QUARTER_MONTHS = {
    1: (1, 2, 3),
    2: (4, 5, 6),
    3: (7, 8, 9),
    4: (10, 11, 12),
}

QUARTER_CODE = {1: "03", 2: "06", 3: "09", 4: "12"}


def quarter_from_month(month):
    return (month - 1) // 3 + 1


def year_quarter_code(year, quarter):
    """Returns CCYYQQ string e.g. 202603 for Q1 2026."""
    return f"{year}{QUARTER_CODE[quarter]}"


def detect_quarter_from_detail(df):
    """
    Infer the reporting quarter from Detail Data pay periods.
    Uses Compensation rows — finds the most common quarter among
    pay period END dates that fall within a single calendar year.
    """
    comp = df[df["Tax Tracking Type"] == "Compensation"].copy()
    comp = comp.dropna(subset=["Pay Period End", "Amount"])
    comp = comp[comp["Amount"] > 0]
    if comp.empty:
        return None, None

    ends = pd.to_datetime(comp["Pay Period End"])
    yq_counts = {}
    for dt in ends:
        q = quarter_from_month(dt.month)
        key = (dt.year, q)
        yq_counts[key] = yq_counts.get(key, 0) + 1

    best = max(yq_counts, key=yq_counts.get)
    return best  # (year, quarter)


def worked_on_12th(comp_by_ssn, ssn, year, month):
    """Return 1 if employee had a Compensation row with Amount>0 whose pay period spans the 12th."""
    target = datetime.date(year, month, 12)
    for start, end in comp_by_ssn.get(ssn, []):
        if start <= target <= end:
            return 1
    return 0


# ---------------------------------------------------------------------------
# CSV field helper
# ---------------------------------------------------------------------------

def csv_field(val, maxlen):
    """
    Sanitize a free-text field for MiUI's delimited format.
    Commas and double-quotes are stripped outright -- MiUI is a
    fixed-format parser that does not support RFC 4180 quoting,
    so any embedded comma would be misread as a field delimiter.
    """
    if not pd.notna(val) or str(val).strip() == "":
        return ""
    s = str(val).strip()
    s = s.replace(",", "").replace('"', "")
    return s[:maxlen]



# ---------------------------------------------------------------------------
# Employer info extraction
# ---------------------------------------------------------------------------

def extract_employer_info(xlsm_path):
    """
    Read employer name, address, and EAN directly from the QB xlsm export.
    Supporting Settings sheet holds the QB legal company info (col6=key, col7=value).
    Detail Data Agency ID column holds the 10-digit SUI account number (7-digit EAN + 000).
    Returns a dict; any field not found is returned as an empty string.
    """
    try:
        detail = pd.read_excel(xlsm_path, sheet_name="Detail Data", header=0)
        ss     = pd.read_excel(xlsm_path, sheet_name="Supporting Settings", header=None)
    except Exception:
        return {}

    kv = {}
    for _, row in ss.iterrows():
        vals = list(row)
        if len(vals) > 6 and pd.notna(vals[6]):
            kv[str(vals[6]).strip()] = str(vals[7]).strip() if len(vals) > 7 and pd.notna(vals[7]) else ""

    # Zip may read as float (49855.0) — strip the decimal
    raw_zip = kv.get("QBLegalZip", "")
    zipcode = raw_zip.split(".")[0] if "." in raw_zip else raw_zip

    info = {
        "employer_name": kv.get("QBLegalCompanyName", ""),
        "addr1":         kv.get("QBLegalAddress1", ""),
        "addr2":         kv.get("QBLegalAddress2", ""),
        "city":          kv.get("QBLegalCity", ""),
        "state":         kv.get("QBLegalState", ""),
        "zipcode":       zipcode,
    }

    # EAN: Agency ID column contains 10-digit values = 7-digit EAN + 3-digit unit suffix
    try:
        agency_ids = detail["Agency ID"].dropna().astype(str)
        sui_ids = agency_ids[agency_ids.str.match(r"^\d{10}$")]
        info["ean"] = sui_ids.iloc[0][:7] if not sui_ids.empty else ""
    except Exception:
        info["ean"] = ""

    return info

# ---------------------------------------------------------------------------
# Conversion logic
# ---------------------------------------------------------------------------

def convert(xlsm_path, output_path, employer_name, ean, addr1, addr2, city, state,
            zipcode, zipext, apportionment, terminating, log_cb):

    log_cb("Reading workbook...")
    try:
        state_report = pd.read_excel(xlsm_path, sheet_name="State Report", header=0)
        detail_data  = pd.read_excel(xlsm_path, sheet_name="Detail Data",  header=0)
    except Exception as e:
        raise RuntimeError(f"Could not read workbook: {e}")

    # Validate required columns
    sr_required = {"SSN", "First Name", "MI", "Last Name", "MI SUI Gross Wages", "Family Owned"}
    dd_required = {"SSN", "Tax Tracking Type", "Amount", "Pay Period Start", "Pay Period End"}
    if not sr_required.issubset(state_report.columns):
        raise RuntimeError(f"State Report sheet is missing columns: {sr_required - set(state_report.columns)}")
    if not dd_required.issubset(detail_data.columns):
        raise RuntimeError(f"Detail Data sheet is missing columns: {dd_required - set(detail_data.columns)}")

    # Detect quarter
    log_cb("Detecting reporting quarter...")
    yq = detect_quarter_from_detail(detail_data)
    if yq is None:
        raise RuntimeError("Could not detect reporting quarter from Detail Data.")
    year, quarter = yq
    yq_code = year_quarter_code(year, quarter)
    months = QUARTER_MONTHS[quarter]
    log_cb(f"Detected: Q{quarter} {year}  ->  Year/Quarter code: {yq_code}")

    # Build 12th-of-month lookup: ssn -> list of (start_date, end_date)
    log_cb("Building 12th-of-month employment data...")
    comp = detail_data[
        (detail_data["Tax Tracking Type"] == "Compensation") &
        (detail_data["Amount"] > 0)
    ].copy()
    comp["Pay Period Start"] = pd.to_datetime(comp["Pay Period Start"])
    comp["Pay Period End"]   = pd.to_datetime(comp["Pay Period End"])

    comp_by_ssn = {}
    for _, row in comp.iterrows():
        ssn   = str(row["SSN"]).strip()
        start = row["Pay Period Start"].date()
        end   = row["Pay Period End"].date()
        comp_by_ssn.setdefault(ssn, []).append((start, end))

    # Filter State Report to employees with wages > 0
    employees = state_report[state_report["MI SUI Gross Wages"] > 0].copy()
    log_cb(f"Employees with wages this quarter: {len(employees)}")

    ean_clean = re.sub(r"[\s\-]", "", str(ean))

    lines = []

    # --- Employer Header Record ---
    # Free-text fields go through csv_field() for comma-safe quoting.
    # Numeric/code fields are written directly.
    re_fields = [
        "RE",
        yq_code,
        ean_clean,
        csv_field(employer_name, 60),
        csv_field(addr1, 45),
        csv_field(addr2, 45),
        csv_field(city, 35),
        str(state).strip().upper()[:2],
        str(zipcode).strip()[:5],
        str(zipext).strip()[:4],
        "Y" if str(apportionment).strip().upper() == "Y" else "N",
        "Y" if str(terminating).strip().upper() == "Y" else "N",
    ]
    lines.append(",".join(re_fields))

    # --- Employee Detail Records ---
    skipped = 0
    written = 0
    for _, emp in employees.iterrows():
        raw_ssn   = str(emp["SSN"]).strip()
        ssn_clean = re.sub(r"[\s\-]", "", raw_ssn)

        if not ssn_clean or ssn_clean == "000000000":
            skipped += 1
            continue

        # Employee names rarely contain commas, but pass through csv_field()
        # for safety (e.g. "Smith, Jr." suffix in a last name field).
        last   = csv_field(emp["Last Name"],  30).upper()
        first  = csv_field(emp["First Name"], 15).upper()
        middle = csv_field(emp["MI"],         30).upper()

        gross     = emp["MI SUI Gross Wages"]
        gross_int = int(round(float(gross) * 100))

        m1 = worked_on_12th(comp_by_ssn, raw_ssn, year, months[0])
        m2 = worked_on_12th(comp_by_ssn, raw_ssn, year, months[1])
        m3 = worked_on_12th(comp_by_ssn, raw_ssn, year, months[2])

        family = str(emp.get("Family Owned", "No")).strip().lower()
        owner  = "Y" if family in ("yes", "y") else "N"

        rw_fields = [
            "RW",
            ean_clean,
            "000",
            yq_code,
            str(m1),
            str(m2),
            str(m3),
            ssn_clean,
            last,
            first,
            middle,
            str(gross_int),
            "N",    # Seasonal indicator
            owner,  # Owner/Officer indicator
            "0",    # Adjustment Reason (original filing)
            "0",    # Obligation Gross Wages (not currently active)
            "N",    # Visa Wage Indicator
        ]
        lines.append(",".join(rw_fields))
        written += 1

    log_cb(f"Writing {written} employee records...")
    if skipped:
        log_cb(f"  WARNING: Skipped {skipped} employee(s) with missing/invalid SSN")

    with open(output_path, "w", newline="\n", encoding="ascii") as f:
        f.write("\n".join(lines) + "\n")

    log_cb(f"\nDone!  Output saved to:\n    {output_path}")
    log_cb(f"\nSummary:")
    log_cb(f"  Quarter:    Q{quarter} {year}  ({yq_code})")
    log_cb(f"  EAN:        {ean_clean}")
    log_cb(f"  Employer:   {employer_name}")
    log_cb(f"  Employees:  {written} records written")
    if skipped:
        log_cb(f"  Skipped:    {skipped} (missing SSN)")


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Michigan MiUI SUI Wage Report Converter")
        self.resizable(False, False)
        self.configure(bg="#f0f0f0")

        self.font_label  = ("Segoe UI", 9)
        self.font_header = ("Segoe UI", 10, "bold")
        self.font_mono   = ("Consolas", 9)

        self._build_ui()
        self._center()

    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"+{(sw-w)//2}+{(sh-h)//2}")

    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}

        # ── File selection ──────────────────────────────────────────────
        file_frame = ttk.LabelFrame(self, text=" Input File ", padding=8)
        file_frame.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 4))

        self.xlsm_var = tk.StringVar()
        ttk.Label(file_frame, text="QuickBooks .xlsm export:").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(file_frame, textvariable=self.xlsm_var, width=50).grid(row=0, column=1, **pad)
        ttk.Button(file_frame, text="Browse...", command=self._browse_input).grid(row=0, column=2, **pad)

        # ── Employer Info ───────────────────────────────────────────────
        emp_frame = ttk.LabelFrame(self, text=" Employer Header (RE Record) ", padding=8)
        emp_frame.grid(row=1, column=0, sticky="ew", padx=12, pady=4)

        fields = [
            ("Employer Name:",    "emp_name",   50, ""),
            ("EAN (7 digits):",   "ean",        10, ""),
            ("Address Line 1:",   "addr1",      45, ""),
            ("Address Line 2:",   "addr2",      45, ""),
            ("City:",             "city",       30, ""),
            ("State (2-letter):", "state_abbr",  4, ""),
            ("Zip Code:",         "zipcode",     6, ""),
            ("Zip Ext (4 dig):",  "zipext",      5, ""),
        ]

        self.vars = {}
        for i, (label, key, width, default) in enumerate(fields):
            ttk.Label(emp_frame, text=label).grid(row=i, column=0, sticky="w", **pad)
            var = tk.StringVar(value=default)
            self.vars[key] = var
            ttk.Entry(emp_frame, textvariable=var, width=width).grid(row=i, column=1, sticky="w", **pad)

        self.apport_var = tk.StringVar(value="N")
        self.term_var   = tk.StringVar(value="N")
        ttk.Label(emp_frame, text="Apportionment Program:").grid(row=len(fields),   column=0, sticky="w", **pad)
        ttk.Label(emp_frame, text="Terminating Business:") .grid(row=len(fields)+1, column=0, sticky="w", **pad)
        ttk.Combobox(emp_frame, textvariable=self.apport_var, values=["N", "Y"], width=4, state="readonly").grid(
            row=len(fields),   column=1, sticky="w", **pad)
        ttk.Combobox(emp_frame, textvariable=self.term_var,   values=["N", "Y"], width=4, state="readonly").grid(
            row=len(fields)+1, column=1, sticky="w", **pad)

        # ── Output File ─────────────────────────────────────────────────
        out_frame = ttk.LabelFrame(self, text=" Output File ", padding=8)
        out_frame.grid(row=2, column=0, sticky="ew", padx=12, pady=4)

        self.out_var = tk.StringVar()
        ttk.Label(out_frame, text="Save CSV to:").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(out_frame, textvariable=self.out_var, width=50).grid(row=0, column=1, **pad)
        ttk.Button(out_frame, text="Browse...", command=self._browse_output).grid(row=0, column=2, **pad)

        # ── Buttons ─────────────────────────────────────────────────────
        btn_frame = ttk.Frame(self)
        btn_frame.grid(row=3, column=0, pady=6)
        self.convert_btn = ttk.Button(btn_frame, text="Convert", command=self._run, width=18)
        self.convert_btn.grid(row=0, column=0, padx=6)
        ttk.Button(btn_frame, text="Clear Log", command=self._clear_log, width=12).grid(row=0, column=1, padx=6)

        # ── Log output ──────────────────────────────────────────────────
        log_frame = ttk.LabelFrame(self, text=" Output Log ", padding=8)
        log_frame.grid(row=4, column=0, sticky="nsew", padx=12, pady=(4, 12))

        self.log_text = tk.Text(log_frame, height=14, width=72, font=self.font_mono,
                                bg="#1e1e1e", fg="#d4d4d4", relief="flat",
                                state="disabled", wrap="word")
        scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll.set)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")

    # ── File dialogs ────────────────────────────────────────────────────

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Select QuickBooks SUI Export",
            filetypes=[("Excel Macro Workbook", "*.xlsm"), ("Excel Workbook", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self.xlsm_var.set(path)
            if not self.out_var.get():
                directory = os.path.dirname(path)
                basename  = os.path.splitext(os.path.basename(path))[0]
                self.out_var.set(os.path.join(directory, f"{basename}_MiUI.csv"))
            self._autofill_employer(path)

    def _autofill_employer(self, path):
        """Attempt to pre-populate employer fields from the selected workbook."""
        info = extract_employer_info(path)
        if not info:
            return
        field_map = {
            "emp_name":   "employer_name",
            "ean":        "ean",
            "addr1":      "addr1",
            "addr2":      "addr2",
            "city":       "city",
            "state_abbr": "state",
            "zipcode":    "zipcode",
        }
        filled = []
        for field_key, info_key in field_map.items():
            val = info.get(info_key, "")
            if val:
                self.vars[field_key].set(val)
                filled.append(field_key)
        if filled:
            self._log("Auto-filled from workbook: " + ", ".join(filled))
            self._log("  Review fields before converting - correct anything that looks wrong.")

    def _browse_output(self):
        initial_dir  = os.path.dirname(self.xlsm_var.get()) if self.xlsm_var.get() else os.path.expanduser("~")
        initial_file = os.path.basename(self.out_var.get()) if self.out_var.get() else "SUI_MiUI.csv"
        path = filedialog.asksaveasfilename(
            title="Save MiUI CSV As",
            initialdir=initial_dir,
            initialfile=initial_file,
            defaultextension=".csv",
            filetypes=[("CSV file", "*.csv"), ("Text file", "*.txt"), ("All files", "*.*")]
        )
        if path:
            self.out_var.set(path)

    # ── Logging ─────────────────────────────────────────────────────────

    def _log(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    # ── Validation & run ─────────────────────────────────────────────────

    def _validate(self):
        if not self.xlsm_var.get() or not os.path.isfile(self.xlsm_var.get()):
            messagebox.showerror("Input Required", "Please select a valid .xlsm input file.")
            return False
        if not self.out_var.get():
            messagebox.showerror("Output Required", "Please specify an output file path.")
            return False
        ean = re.sub(r"[\s\-]", "", self.vars["ean"].get())
        if not ean.isdigit() or len(ean) != 7:
            messagebox.showerror("Invalid EAN", "EAN must be exactly 7 digits (no hyphens or spaces).")
            return False
        if not self.vars["emp_name"].get().strip():
            messagebox.showerror("Missing Field", "Employer Name is required.")
            return False
        if not self.vars["city"].get().strip():
            messagebox.showerror("Missing Field", "City is required.")
            return False
        zipcode = self.vars["zipcode"].get().strip()
        if not zipcode.isdigit() or len(zipcode) != 5:
            messagebox.showerror("Invalid Zip", "Zip Code must be exactly 5 digits.")
            return False
        return True

    def _run(self):
        if not self._validate():
            return
        self.convert_btn.configure(state="disabled", text="Converting...")
        self._clear_log()
        self._log("=" * 60)
        self._log("  Michigan MiUI SUI Wage Report Converter")
        self._log("=" * 60)

        def worker():
            try:
                convert(
                    xlsm_path     = self.xlsm_var.get(),
                    output_path   = self.out_var.get(),
                    employer_name = self.vars["emp_name"].get().strip(),
                    ean           = self.vars["ean"].get().strip(),
                    addr1         = self.vars["addr1"].get().strip(),
                    addr2         = self.vars["addr2"].get().strip(),
                    city          = self.vars["city"].get().strip(),
                    state         = self.vars["state_abbr"].get().strip().upper(),
                    zipcode       = self.vars["zipcode"].get().strip(),
                    zipext        = self.vars["zipext"].get().strip(),
                    apportionment = self.apport_var.get(),
                    terminating   = self.term_var.get(),
                    log_cb        = lambda msg: self.after(0, self._log, msg),
                )
            except Exception as e:
                self.after(0, self._log, f"\nError: {e}")
                self.after(0, messagebox.showerror, "Conversion Failed", str(e))
            finally:
                self.after(0, self.convert_btn.configure, {"state": "normal", "text": "Convert"})

        threading.Thread(target=worker, daemon=True).start()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app = App()
    app.mainloop()
