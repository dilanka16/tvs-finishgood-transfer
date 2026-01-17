import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pyodbc
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import pandas as pd
from tkcalendar import DateEntry
import threading

# ---------- SQL SERVER ----------
server = "192.168.5.253"
database = "GTDBSLTL"
username = "sa"
password = "abcd@abcd"

# ---------- GOOGLE SHEETS ----------
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1PnFKuVZJKMY7oNrxrvTrbEFc_pp8QpTqUHhbANjtl6U/edit"
SERVICE_ACCOUNT_FILE = "C:/xampp/htdocs/curing-time-386c0cb6af9e.json"

scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]
creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, scope)
client = gspread.authorize(creds)
spreadsheet = client.open_by_url(GOOGLE_SHEET_URL)

scanned_barcodes = set()

# ---------- SQL FETCH ----------
def fetch_barcode_details(cur, barcode):
    try:
        cur.execute("""
            SELECT CreateDate,
                   CONCAT(tyresize,'',config,'',brand),
                   stencilno,
                   barcode,
                   DispatchStatus
            FROM tbqualitycontrol WHERE barcode = ?
        """, barcode)
        row = cur.fetchone()
        if not row:
            return None
        if row[4] == "Yes":
            return "DISPATCHED"
        return row
    except Exception as e:
        messagebox.showerror("SQL Error", str(e))
        return None

# ---------- PALLET ----------
def get_today_pallet():
    sheet = spreadsheet.worksheet("FGST")
    rows = sheet.get_all_values()[1:]
    today = datetime.now().strftime("%Y-%m-%d")
    pallets = [r[6] for r in rows if len(r) >= 7 and r[5] == today]
    if not pallets:
        return "Pallet 1"
    nums = [int(p.split()[-1]) for p in pallets]
    return f"Pallet {max(nums)+1}"

# ---------- SCAN WINDOW ----------
def open_scan_page():
    win = tk.Toplevel(root)
    win.title("Transfer FGS Tyres")
    win.geometry("1100x700")

    tk.Label(win, text="FGS Barcode Scanner", font=("Arial",18,"bold")).pack(pady=10)
    live_lbl = tk.Label(win, text="0 Ready to Scan", font=("Arial",14), fg="green")
    live_lbl.pack(anchor="w", padx=20)

    entry = tk.Text(win, height=5, font=("Arial",14))
    entry.pack(fill="x", padx=20)
    entry.focus()

    cols = ("cd","sd","name","stencil","barcode","date","pallet","del")
    tree = ttk.Treeview(win, columns=cols, show="headings", height=15)
    heads = ["Create Date","Scanned Time","Tyre Name","Stencil No",
             "Barcode","Date","Pallet","Delete"]
    widths = [150,170,220,120,170,110,90,70]
    for c,h,w in zip(cols,heads,widths):
        tree.heading(c,text=h)
        tree.column(c,width=w,anchor="center")
    tree.pack(fill="both", expand=True, pady=10)

    today = datetime.now().strftime("%Y-%m-%d")
    pallet = get_today_pallet()
    scanned_rows = []

    # ---------- PERSISTENT SQL CONNECTION ----------
    conn = pyodbc.connect(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={server};DATABASE={database};"
        f"UID={username};PWD={password}"
    )
    cur = conn.cursor()

    # ---------- LIVE COUNT ----------
    def update_live_count(event=None):
        lines = entry.get("1.0", tk.END).strip().splitlines()
        unique_barcodes = [b for b in lines if b and b not in scanned_barcodes]
        live_lbl.config(text=f"{len(unique_barcodes)} Ready to Scan")
    entry.bind("<KeyRelease>", update_live_count)
    update_live_count()

    # ---------- SCAN FUNCTION ----------
    def scan():
        lines = entry.get("1.0", tk.END).strip().splitlines()
        new_rows = []
        for b in lines:
            if b in scanned_barcodes:
                continue
            d = fetch_barcode_details(cur, b)
            if not d:
                continue
            if d == "DISPATCHED":
                messagebox.showwarning("Already Dispatched", b)
                continue
            row = [
                str(d[0]),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                d[1],
                d[2],
                d[3],
                today,
                pallet,
                "❌"
            ]
            scanned_rows.append(row)
            scanned_barcodes.add(b)
            new_rows.append(row)

        for row in new_rows:
            tree.insert("", tk.END, values=row)

        entry.delete("1.0", tk.END)
        update_live_count()
        update_count()

    # ---------- SAVE FUNCTION ----------
    def save():
        if not scanned_rows:
            return
        sheet = spreadsheet.worksheet("FGST")
        batch_rows = [r[:-1] for r in scanned_rows]
        sheet.append_rows(batch_rows)
        scanned_rows.clear()
        scanned_barcodes.clear()
        update_count()
        update_pallet_summary()
        win.destroy()
        conn.close()

    tk.Button(win, text="SCAN", font=("Arial",14),
              bg="#0078d7", fg="white", command=scan).pack(side=tk.LEFT, padx=20, pady=10)
    tk.Button(win, text="OK", font=("Arial",14),
              bg="green", fg="white", command=save).pack(side=tk.LEFT, padx=10)

# ---------- DATE RANGE EXPORT ----------
def export_by_date_range():
    win = tk.Toplevel(root)
    win.title("Select Date Range")
    win.geometry("400x250")

    tk.Label(win, text="From Date").pack(pady=5)
    from_e = DateEntry(win, date_pattern='yyyy-mm-dd')
    from_e.pack()

    tk.Label(win, text="To Date").pack(pady=5)
    to_e = DateEntry(win, date_pattern='yyyy-mm-dd')
    to_e.pack()

    def download():
        f, t = from_e.get_date().strftime("%Y-%m-%d"), to_e.get_date().strftime("%Y-%m-%d")
        sheet = spreadsheet.worksheet("FGST")
        rows = sheet.get_all_values()[1:]
        data = [r[:5] for r in rows if len(r)>=5 and f <= r[1][:10] <= t]
        if not data:
            messagebox.showinfo("No Data", "No records found")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files","*.xlsx")]
        )
        if not path:
            return

        df = pd.DataFrame(data, columns=[
            "Create Date","Scanned Date & Time",
            "Tyre Name","Stencil No","Barcode"
        ])
        df.to_excel(path, index=False)
        messagebox.showinfo("Success", f"{len(data)} records exported")

    tk.Button(win, text="Download Excel",
              font=("Arial",14),
              bg="#9c27b0", fg="white",
              command=download).pack(pady=20)

# ---------- SEARCH ----------
def search_stencil():
    stencil = search_entry.get().strip()
    sheet = spreadsheet.worksheet("FGST")
    rows = sheet.get_all_values()[1:]
    found = [r for r in rows if len(r)>=5 and r[3]==stencil]
    if not found:
        messagebox.showinfo("Not Found","Stencil No not found")
        return

    win = tk.Toplevel(root)
    win.title(f"Stencil {stencil} Details")
    win.geometry("900x400")

    cols = ("create","scanned","name","stencil","barcode","date","pallet")
    tree = ttk.Treeview(win, columns=cols, show="headings")
    headers = ["Create Date","Scanned Time","Tyre Name","Stencil No","Barcode","Date","Pallet"]
    widths = [150,150,200,120,150,100,100]

    for c,h,w in zip(cols,headers,widths):
        tree.heading(c,text=h)
        tree.column(c,width=w,anchor="center")

    for r in found:
        tree.insert("", tk.END, values=r[:7])

    tree.pack(fill="both", expand=True, pady=10)

def search_by_date():
    date = date_picker.get_date().strftime("%Y-%m-%d")
    sheet = spreadsheet.worksheet("FGST")
    rows = sheet.get_all_values()[1:]
    found = [r for r in rows if len(r)>=5 and r[1].startswith(date)]
    messagebox.showinfo("Result", f"{len(found)} scanned on {date}")

# ---------- MAIN PAGE ----------
root = tk.Tk()
root.title("FGS Transfer System")
root.geometry("900x700")

tk.Label(root, text="FGS Tyre Transfer", font=("Arial",24,"bold")).pack(pady=20)

search_frame = tk.Frame(root)
search_frame.pack(pady=10)

search_entry = tk.Entry(search_frame, font=("Arial",14), width=25)
search_entry.pack(side=tk.LEFT, padx=5)
tk.Button(search_frame, text="Find Stencil No",
          bg="#ff9800", command=search_stencil).pack(side=tk.LEFT, padx=5)

date_picker = DateEntry(search_frame, date_pattern='yyyy-mm-dd', font=("Arial",14))
date_picker.pack(side=tk.LEFT, padx=5)
tk.Button(search_frame, text="Search by Date",
          bg="#2196f3", fg="white",
          command=search_by_date).pack(side=tk.LEFT, padx=5)

tk.Button(search_frame,
          text="Select Date Range & Download Excel",
          bg="#9c27b0", fg="white",
          command=export_by_date_range).pack(side=tk.LEFT, padx=5)

# ---------- PALLET SUMMARY BUTTON ----------
def filter_summary_by_date():
    date = date_picker.get_date().strftime("%Y-%m-%d")
    update_pallet_summary(date)

tk.Button(search_frame,
          text="Show Pallet Summary for Selected Date",
          bg="#4caf50", fg="white",
          command=filter_summary_by_date).pack(side=tk.LEFT, padx=5)

# ---------- TRANSFER BUTTON ----------
btn_frame = tk.Frame(root)
btn_frame.pack(pady=20)

tk.Button(
    btn_frame,
    text="Transfer FGS Tyres",
    font=("Arial",20,"bold"),
    width=18,
    height=3,
    bg="#28a745",
    fg="white",
    command=open_scan_page
).pack()

count_label = tk.Label(btn_frame, text="0 Scanned", font=("Arial",14))
count_label.pack(pady=10)

def update_count():
    count_label.config(text=f"{len(scanned_barcodes)} Scanned")

# ---------- DAY-WISE PALLET SUMMARY ----------
summary_cols = ("date","pallet","qty")
summary_tree = ttk.Treeview(root, columns=summary_cols, show="headings", height=8)
summary_tree.heading("date", text="Date")
summary_tree.heading("pallet", text="Pallet No")
summary_tree.heading("qty", text="Qty")
summary_tree.column("date", width=120, anchor="center")
summary_tree.column("pallet", width=120, anchor="center")
summary_tree.column("qty", width=80, anchor="center")
summary_tree.pack(fill="x", padx=20, pady=10)

def update_pallet_summary(selected_date=None):
    summary_tree.delete(*summary_tree.get_children())
    sheet = spreadsheet.worksheet("FGST")
    rows = sheet.get_all_values()[1:]

    if selected_date is None:
        selected_date = datetime.now().strftime("%Y-%m-%d")

    pallet_dict = {}
    total_qty = 0
    for r in rows:
        if len(r) >= 7 and r[5] == selected_date:
            pallet_no = r[6]
            pallet_dict[pallet_no] = pallet_dict.get(pallet_no, 0) + 1
            total_qty += 1

    for pallet, qty in sorted(pallet_dict.items()):
        summary_tree.insert("", tk.END, values=(selected_date, pallet, qty))
    summary_tree.insert("", tk.END, values=(selected_date, "Total", total_qty))

update_pallet_summary()
root.mainloop()
