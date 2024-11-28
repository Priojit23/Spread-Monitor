import MetaTrader5 as mt5
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import multiprocessing
import pandas as pd
import time

# MT5 Installation Paths
mt5_paths = {
    "ECN": r"C:\Program Files\MetaTrader 5 EXNESS\terminal64.exe",
    "RAW": r"C:\Program Files\Tickmill MT5 Terminal\terminal64.exe",
  
    "STD": r"C:\Program Files\FundedNext MT5 Terminal\terminal64.exe",
"Classic": r"C:\Program Files\MetaTrader 5 IC Markets Global\terminal64.exe"
}

# Account Details
accounts = {
    "Classic": {
        "login": 55683316,
        "password": "b2%q(1Dhn!<?",
        "server": "Tickmill-Live"
    },
    "RAW": {
        "login": 24088398,
        "password": "Investor5ers.",
        "server": "FivePercentOnline-Real"
    },
    "STD": {
        "login": 13109044,
        "password": "ubmAX18##",
        "server": "FundedNext-Server 2"
    },
    "ECN": {
        "login": 1510140547,
        "password": "6y*8P1V9?",
        "server": "FTMO-Demo"
    }
}

# Function to fetch spreads in a separate process
def fetch_spreads(account_name, account_details, mt5_path, output_queue):
    if not mt5.initialize(path=mt5_path):
        output_queue.put((account_name, f"Failed to initialize MT5 for {account_name}: {mt5.last_error()}"))
        return

    if not mt5.login(login=account_details["login"], password=account_details["password"], server=account_details["server"]):
        output_queue.put((account_name, f"Failed to log in to {account_name}: {mt5.last_error()}"))
        mt5.shutdown()
        return

    symbols = mt5.symbols_get()
    symbol_names = [symbol.name for symbol in symbols]

    for symbol in symbols:
        if not mt5.symbol_info(symbol.name).visible:
            mt5.symbol_select(symbol.name, True)

    output_queue.put((account_name, f"Successfully logged in to {account_name}"))

    try:
        while True:
            spreads = {}
            for symbol_name in symbol_names:
                tick = mt5.symbol_info_tick(symbol_name)
                symbol_info = mt5.symbol_info(symbol_name)
                if tick and symbol_info:
                    bid = tick.bid
                    ask = tick.ask
                    point = symbol_info.point
                    if point > 0:
                        spread = round((ask - bid) / point, 1)
                        # Remove suffixes (.std or .raw)
                        clean_symbol_name = symbol_name.split(".")[0]
                        spreads[clean_symbol_name] = spread
            output_queue.put((account_name, spreads))
            time.sleep(0.01)
    except Exception as e:
        output_queue.put((account_name, f"Error: {e}"))
    finally:
        mt5.shutdown()

# Function to apply search filtering
def apply_filter():
    query = search_var.get().lower()
    tree.delete(*tree.get_children())  # Clear existing rows in the Treeview

    # Filter combined data and update the Treeview
    for symbol, spreads in combined_data.items():
        if query in symbol.lower():  # Check if the query matches the symbol
            tree.insert(
                "",
                tk.END,
                values=(
                    symbol,
                    spreads.get("ECN", "N/A"),
                    spreads.get("RAW", "N/A"),
                    
                    spreads.get("STD", "N/A"),
                spreads.get("Classic", "N/A")
)
            )

# Function to update the consolidated Treeview
def update_gui_consolidated():
    tree.delete(*tree.get_children())
    for symbol, spreads in combined_data.items():
        tree.insert(
            "",
            tk.END,
            values=(
                symbol,
                spreads.get("ECN", "N/A"),
                spreads.get("RAW", "N/A"),
                
                spreads.get("STD", "N/A"),
            spreads.get("Classic", "N/A")
)
        )
    apply_filter()  # Ensure filtering is applied after each update

# Function to download the consolidated data into an Excel file
def download_spread_data():
    if not combined_data:
        print("No data available for download.")
        return

    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
        title="Save Consolidated Spread Data"
    )

    if not file_path:
        return

    df = pd.DataFrame.from_dict(combined_data, orient="index")
    df.reset_index(inplace=True)
    df.columns = ["Symbol", "ECN", "RAW", "STD", "Classic"]

    df.to_excel(file_path, index=False)
    print(f"Consolidated spread data saved to {file_path}")

# Main function to create the GUI
def main_consolidated():
    global root, tree, combined_data, data_queues, search_var
    combined_data = {}
    data_queues = {}

    root = tk.Tk()
    root.title("Consolidated Spread Monitor with Search")

    # Create search bar
    search_var = tk.StringVar()
    search_var.trace_add("write", lambda *args: apply_filter())  # Update Treeview on search
    search_entry = ttk.Entry(root, textvariable=search_var)
    search_entry.pack(fill=tk.X, padx=5, pady=5)

    # Create the Treeview
    columns = ("Symbol", "ECN", "RAW", "STD", "Classic")
    tree = ttk.Treeview(root, columns=columns, show="headings", height=25)
    tree.heading("Symbol", text="Symbol")
    tree.heading("ECN", text="ECN")
    tree.heading("RAW", text="RAW")
    tree.heading("Classic", text="Classic")
    tree.heading("STD", text="STD")
    tree.pack(expand=True, fill=tk.BOTH)

    # Configure Treeview style
    style = ttk.Style()
    style.configure("Treeview.Heading", anchor="center")
    style.configure("Treeview", rowheight=25)
    tree.column("Symbol", anchor="center", width=150)
    tree.column("ECN", anchor="center", width=100)
    tree.column("RAW", anchor="center", width=100)
    tree.column("Classic", anchor="center", width=100)
    tree.column("STD", anchor="center", width=100)

    # Add a Download button
    download_button = tk.Button(root, text="Download Spread Data", command=download_spread_data)
    download_button.pack(pady=10)

    processes = []

    # Start data collection processes for each account
    for account_name, account_details in accounts.items():
        data_queues[account_name] = multiprocessing.Queue()
        p = multiprocessing.Process(
            target=fetch_spreads,
            args=(account_name, account_details, mt5_paths[account_name], data_queues[account_name])
        )
        p.start()
        processes.append(p)

    # Start updating the Treeview
    def fetch_and_update():
        for account_name, queue in data_queues.items():
            while not queue.empty():
                account_name, data = queue.get()
                if isinstance(data, dict):
                    for symbol, spread in data.items():
                        if symbol not in combined_data:
                            combined_data[symbol] = {}
                        combined_data[symbol][account_name] = spread

    def update_loop():
        fetch_and_update()
        update_gui_consolidated()
        root.after(1000, update_loop)

    update_loop()
    root.mainloop()

    for p in processes:
        p.terminate()
        p.join()

if __name__ == "__main__":
    multiprocessing.freeze_support()
    main_consolidated()
# Spread-Monitor
