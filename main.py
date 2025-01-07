import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import requests
from bs4 import BeautifulSoup
import yfinance as yf
import os
from openpyxl import Workbook

# File to store tickers between sessions
TICKERS_FILE = "tickers.txt"

# Store just the tickers in the order they were added (no duplicates)
tickers_list = []

def load_tickers():
    """Load tickers from TICKERS_FILE. If file doesn't exist, create one."""
    if os.path.exists(TICKERS_FILE):
        with open(TICKERS_FILE, "r") as f:
            lines = f.read().strip().splitlines()
            for line in lines:
                ticker = line.strip().upper()
                if ticker and ticker not in tickers_list:
                    tickers_list.append(ticker)

def save_tickers():
    """Save the current tickers to TICKERS_FILE."""
    with open(TICKERS_FILE, "w") as f:
        for t in tickers_list:
            f.write(t + "\n")

def scrape_finviz_and_fill(ticker: str) -> dict:
    """Fetch information about the given ticker from Finviz and yfinance."""
    url = f"https://finviz.com/quote.ashx?t={ticker}&p=d"
    headers = {"User-Agent": "Mozilla/5.0"}
    response = requests.get(url, headers=headers)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")

    desired_fields = {
        "Company": None,
        "Sector": None,
        "Price": None,
        "Change 5Y": None,
        "Dividends": "Dividend Est.",
        "Dividend TTM": "Dividend TTM",
        "EPS": "EPS (ttm)",
        "EPS Next Y, %": "EPS next Y",
        "EPS Next 5Y, %": "EPS next 5Y",
        "Revenue": "Sales",
        "Revenue 5Y growth": "Sales past 5Y",
        "Oper. Income": "Oper. Margin",
        "Net Income": "Profit Margin",
        "ROA": "ROA",
        "ROE": "ROE",
        "ROI": "ROI",
        "P/E": "P/E",
        "P/S": "P/S",
        "P/B": "P/B"
    }

    snapshot_table = soup.find("table", class_="snapshot-table2")
    field_values = {}
    if snapshot_table:
        tds = snapshot_table.find_all("td")
        for i in range(0, len(tds), 2):
            label = tds[i].get_text(strip=True)
            value = tds[i + 1].get_text(strip=True) if (i + 1) < len(tds) else "N/A"
            field_values[label] = value

    # Company name
    company_tag = soup.find("h2", class_="quote-header_ticker-wrapper_company")
    company_name = company_tag.get_text(strip=True) if company_tag else "N/A"

    # Sector
    sector_div = soup.find("div", class_="flex space-x-0.5 overflow-hidden")
    sector = "N/A"
    if sector_div:
        sector_links = sector_div.find_all("a")
        if len(sector_links) >= 1:
            sector = sector_links[0].get_text(strip=True)

    # Price
    price_tag = soup.find("strong", class_="quote-price_wrapper_price")
    price = price_tag.get_text(strip=True) if price_tag else "N/A"

    # Prepare result dict with some direct fields
    result = {
        "Ticker": ticker,
        "Company": company_name if company_name else "N/A",
        "Sector": sector,
        "Price": price
    }

    # Fill in Finviz-based fields
    for field_name, finviz_label in desired_fields.items():
        if field_name in ["Company", "Sector", "Price"]:
            continue
        if finviz_label is None:
            result[field_name] = "N/A"
        else:
            result[field_name] = field_values.get(finviz_label, "N/A")

    # Use yfinance to fetch balance sheet data
    yf_ticker = yf.Ticker(ticker)
    try:
        bal_sheet = yf_ticker.balance_sheet
        if not bal_sheet.empty:
            if "Total Assets" in bal_sheet.index:
                total_assets = bal_sheet.loc["Total Assets"].iloc[0]
                if isinstance(total_assets, (int, float)):
                    result["Total Assets"] = f"{total_assets:,}"
                else:
                    result["Total Assets"] = str(total_assets)
            else:
                result["Total Assets"] = "N/A"

            if "Total Liabilities Net Minority Interest" in bal_sheet.index:
                total_liab = bal_sheet.loc["Total Liabilities Net Minority Interest"].iloc[0]
                if isinstance(total_liab, (int, float)):
                    result["Total Liabilities"] = f"{total_liab:,}"
                else:
                    result["Total Liabilities"] = str(total_liab)
            else:
                result["Total Liabilities"] = "N/A"
        else:
            result["Total Assets"] = "N/A"
            result["Total Liabilities"] = "N/A"
    except Exception:
        result["Total Assets"] = "N/A"
        result["Total Liabilities"] = "N/A"

    # Compute 5-Year Price Change
    try:
        hist_5y = yf_ticker.history(period='5y')
        if not hist_5y.empty:
            start_price = hist_5y['Close'].iloc[0]
            end_price = hist_5y['Close'].iloc[-1]
            change_5y_val = ((end_price - start_price) / start_price) * 100
            result["Change 5Y"] = f"{change_5y_val:.2f}%"
        else:
            result["Change 5Y"] = "N/A"
    except Exception:
        result["Change 5Y"] = "N/A"

    return result

def fetch_data():
    """Add the entered ticker to the list if not already present, and refresh data."""
    ticker = ticker_entry.get().strip().upper()
    if not ticker:
        messagebox.showwarning("Warning", "Please enter a ticker.")
        return

    if ticker in tickers_list:
        messagebox.showinfo("Info", f"{ticker} is already in the list.")
        ticker_entry.delete(0, tk.END)
        return

    tickers_list.append(ticker)
    ticker_entry.delete(0, tk.END)
    save_tickers()
    refresh_table()

def refresh_table():
    """Fetch data for all tickers in the list and display it in the table."""
    # Clear existing rows
    for row in table.get_children():
        table.delete(row)

    for ticker in tickers_list:
        try:
            data = scrape_finviz_and_fill(ticker)
        except requests.exceptions.HTTPError as e:
            messagebox.showerror("Error", f"HTTP Error fetching {ticker}: {e}")
            continue
        except Exception as e:
            messagebox.showerror("Error", f"Error fetching {ticker}: {e}")
            continue

        table.insert("", "end", values=(
            data.get("Ticker", "N/A"),
            data.get("Company", "N/A"),
            data.get("Sector", "N/A"),
            data.get("Price", "N/A"),
            data.get("Change 5Y", "N/A"),
            data.get("Dividends", "N/A"),
            data.get("Dividend TTM", "N/A"),
            data.get("EPS", "N/A"),
            data.get("EPS Next Y, %", "N/A"),
            data.get("EPS Next 5Y, %", "N/A"),
            data.get("Revenue", "N/A"),
            data.get("Revenue 5Y growth", "N/A"),
            data.get("Oper. Income", "N/A"),
            data.get("Net Income", "N/A"),
            data.get("ROA", "N/A"),
            data.get("ROE", "N/A"),
            data.get("ROI", "N/A"),
            data.get("P/E", "N/A"),
            data.get("P/S", "N/A"),
            data.get("P/B", "N/A"),
            data.get("Total Assets", "N/A"),
            data.get("Total Liabilities", "N/A")
        ))

def clear_tickers():
    """Clear the ticker list and update the table."""
    tickers_list.clear()
    save_tickers()
    refresh_table()
    # Also clear the side panel
    show_ticker_details({})

def export_to_xlsx():
    """Fetch fresh data for all tickers and export to an Excel file."""
    if not tickers_list:
        messagebox.showinfo("Info", "No tickers to export.")
        return

    # Re-fetch data for all tickers to ensure fresh data
    records = []
    for ticker in tickers_list:
        try:
            data = scrape_finviz_and_fill(ticker)
            records.append(data)
        except:
            pass

    if not records:
        messagebox.showinfo("Info", "No data to export.")
        return

    wb = Workbook()
    ws = wb.active

    # Columns defined in the same order as the table
    columns = (
        "Ticker", "Company", "Sector", "Price", "Change 5Y", "Dividends", "Dividend TTM",
        "EPS", "EPS Next Y, %", "EPS Next 5Y, %", "Revenue", "Revenue 5Y growth",
        "Oper. Income", "Net Income", "ROA", "ROE", "ROI", "P/E", "P/S", "P/B",
        "Total Assets", "Total Liabilities"
    )

    ws.append(columns)  # header row
    for r in records:
        row = (
            r.get("Ticker", "N/A"),
            r.get("Company", "N/A"),
            r.get("Sector", "N/A"),
            r.get("Price", "N/A"),
            r.get("Change 5Y", "N/A"),
            r.get("Dividends", "N/A"),
            r.get("Dividend TTM", "N/A"),
            r.get("EPS", "N/A"),
            r.get("EPS Next Y, %", "N/A"),
            r.get("EPS Next 5Y, %", "N/A"),
            r.get("Revenue", "N/A"),
            r.get("Revenue 5Y growth", "N/A"),
            r.get("Oper. Income", "N/A"),
            r.get("Net Income", "N/A"),
            r.get("ROA", "N/A"),
            r.get("ROE", "N/A"),
            r.get("ROI", "N/A"),
            r.get("P/E", "N/A"),
            r.get("P/S", "N/A"),
            r.get("P/B", "N/A"),
            r.get("Total Assets", "N/A"),
            r.get("Total Liabilities", "N/A")
        )
        ws.append(row)

    # Prompt user to save file
    filepath = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if filepath:
        wb.save(filepath)
        messagebox.showinfo("Info", f"Data exported to {filepath}")

# ------------- Categories -------------
BASIC_INFO_KEYS = ["Ticker", "Company", "Sector", "Price"]
PERFORMANCE_KEYS = ["Change 5Y", "Revenue 5Y growth", "EPS Next Y, %", "EPS Next 5Y, %", "Dividends", "Dividend TTM"]
PROFITABILITY_KEYS = ["Revenue", "Oper. Income", "Net Income", "ROA", "ROE", "ROI"]
VALUATION_KEYS = ["EPS", "P/E", "P/S", "P/B"]
BALANCE_SHEET_KEYS = ["Total Assets", "Total Liabilities"]

def show_ticker_details(data: dict):
    """
    Display the currently selected ticker's details on the right panel,
    laid out as follows:
      Row 1 (side by side): Basic Info, Performance
      Row 2 (side by side): Profitability, Valuation
      Row 3 (full-width):   Balance Sheet
    """
    # Clear the details_frame
    for widget in details_frame.winfo_children():
        widget.destroy()

    # If we have no data (e.g., clearing tickers), show a placeholder.
    if not data:
        ttk.Label(details_frame, text="No ticker selected.", foreground="gray").pack(anchor="w", padx=5, pady=5)
        return

    # Helper: display a category block (title + fields) inside some parent frame
    def display_category_in_frame(parent, title, fields):
        cat_frame = ttk.Frame(parent)
        cat_frame.pack(side="left", fill="both", expand=True, padx=10)

        lbl_title = ttk.Label(cat_frame, text=title, font=("TkDefaultFont", 10, "bold"))
        lbl_title.pack(anchor="w", pady=(0, 5))

        for key in fields:
            val = data.get(key, "N/A")
            line = f"{key}: {val}"
            lbl_line = ttk.Label(cat_frame, text=line)
            lbl_line.pack(anchor="w")

    # --- Row 1: Basic Info (left), Performance (right) ---
    row_frame_1 = ttk.Frame(details_frame)
    row_frame_1.pack(anchor="w", fill="x", pady=10)

    display_category_in_frame(row_frame_1, "Basic Info", BASIC_INFO_KEYS)
    display_category_in_frame(row_frame_1, "Performance", PERFORMANCE_KEYS)

    # --- Row 2: Profitability (left), Valuation (right) ---
    row_frame_2 = ttk.Frame(details_frame)
    row_frame_2.pack(anchor="w", fill="x", pady=10)

    display_category_in_frame(row_frame_2, "Profitability", PROFITABILITY_KEYS)
    display_category_in_frame(row_frame_2, "Valuation", VALUATION_KEYS)

    # --- Row 3: Balance Sheet (full width) ---
    row_frame_3 = ttk.Frame(details_frame)
    row_frame_3.pack(anchor="w", fill="x", pady=10)

    # Just put the "Balance Sheet" category alone in row_frame_3
    display_category_in_frame(row_frame_3, "Balance Sheet", BALANCE_SHEET_KEYS)

def on_table_select(event):
    """
    When the user clicks a row in the table, display its detailed info on the right,
    then resize the window to fit new contents.
    """
    selected_items = table.selection()
    if not selected_items:
        return
    item_id = selected_items[0]
    values = table.item(item_id, "values")
    # Make a dict {column_name: value, ...}
    selected_data = dict(zip(columns, values))

    # Show the details
    show_ticker_details(selected_data)

    # Auto-resize the window to fit new content
    root.update_idletasks()
    root.geometry("")  # "" tells Tkinter to resize to fit content

# ------------------- Main Application -------------------
root = tk.Tk()
root.title("Ticker Data Fetcher")
root.geometry("900x500")  # Initial size; it will auto-resize upon selection

# Allow the root window to expand in both directions
root.rowconfigure(0, weight=1)
root.columnconfigure(0, weight=1)

# Menu bar
menubar = tk.Menu(root)
options_menu = tk.Menu(menubar, tearoff=0)
options_menu.add_command(label="Clear Tickers", command=clear_tickers)
options_menu.add_command(label="Export to XLSX", command=export_to_xlsx)
menubar.add_cascade(label="Options", menu=options_menu)
root.config(menu=menubar)

# Main frame (uses grid so it can expand in both directions)
main_frame = ttk.Frame(root, padding="10")
main_frame.grid(row=0, column=0, sticky="nsew")

# We want two columns in main_frame:
# - Left side: input + table
# - Right side: details panel
main_frame.rowconfigure(1, weight=1)
main_frame.columnconfigure(0, weight=1)  # table area can expand
main_frame.columnconfigure(1, weight=0)  # details panel has a fixed width

# Input section (top row, spanning both columns)
input_frame = ttk.Frame(main_frame)
input_frame.grid(row=0, column=0, sticky="ew", pady=10, columnspan=2)

ttk.Label(input_frame, text="Enter Ticker:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
ticker_entry = ttk.Entry(input_frame, width=20)
ticker_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

fetch_button = ttk.Button(input_frame, text="Fetch Data", command=fetch_data)
fetch_button.grid(row=0, column=2, padx=5, pady=5, sticky="w")

# Table frame (left side)
table_frame = ttk.Frame(main_frame)
table_frame.grid(row=1, column=0, sticky="nsew")

table_frame.rowconfigure(0, weight=1)
table_frame.columnconfigure(0, weight=1)

y_scroll = ttk.Scrollbar(table_frame, orient="vertical")
y_scroll.grid(row=0, column=1, sticky="ns")

x_scroll = ttk.Scrollbar(table_frame, orient="horizontal")
x_scroll.grid(row=1, column=0, sticky="ew")

columns = (
    "Ticker", "Company", "Sector", "Price", "Change 5Y", "Dividends", "Dividend TTM",
    "EPS", "EPS Next Y, %", "EPS Next 5Y, %", "Revenue", "Revenue 5Y growth",
    "Oper. Income", "Net Income", "ROA", "ROE", "ROI", "P/E", "P/S", "P/B",
    "Total Assets", "Total Liabilities"
)

table = ttk.Treeview(
    table_frame,
    columns=columns,
    show="headings",
    yscrollcommand=y_scroll.set,
    xscrollcommand=x_scroll.set
)
table.grid(row=0, column=0, sticky="nsew")
y_scroll.config(command=table.yview)
x_scroll.config(command=table.xview)

# Set heading & column widths
for col in columns:
    table.heading(col, text=col)
    # Large width so horizontal scrolling is required.
    table.column(col, width=150, stretch=False)

# Bind the row selection event
table.bind("<<TreeviewSelect>>", on_table_select)

# Details frame (right side)
details_frame = ttk.Frame(main_frame, padding="10", borderwidth=2, relief="groove")
details_frame.grid(row=1, column=1, sticky="ns", padx=(15,0))

# Just a placeholder label initially
ttk.Label(details_frame, text="No ticker selected yet.", foreground="gray").pack(anchor="w", padx=5, pady=5)

# Load previously stored tickers
load_tickers()

# Ensure AAPL is present by default
if "AAPL" not in tickers_list:
    tickers_list.append("AAPL")
    save_tickers()

refresh_table()
root.mainloop()
