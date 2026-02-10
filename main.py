import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime
import os

# ---------------- FILES ----------------
PENDING_FILE = "pending_orders.xlsx"
EXECUTED_FILE = "executed_orders.xlsx"

COLUMNS = [
    "Order Id", "Party", "Item", "Qty",
    "Price", "Total", "Order Date", "Order Time"
]

PARTIES = ["Aman Traders", "Ratan Agency", "Om Distributors","Sai Traders", "MUKESH WHOLSALER", "VU AGENCY"]
ITEMS = {
    "CORN": 5,
    "KURKURE": 20,
    "KITKAT": 10,
    "BISCUIT": 8,
    "SOFT DRINK":25,
    "CADBURY":7,
    "WHEAT":45,
    "HAIROIL":30

}

# ---------------- DATA FUNCTIONS ----------------
def load_data(file):
    if not os.path.exists(file):
        df = pd.DataFrame(columns=COLUMNS)
        df.to_excel(file, index=False)
    else:
        df = pd.read_excel(file)

    df = df.dropna(how="all").fillna("")
    return df


def save_data(df, file):
    df.to_excel(file, index=False)


# ---------------- UI APP ----------------
root = tk.Tk()
root.title("Order Management System")
root.geometry("1200x700")
root.configure(bg="#f2f4f7")

style = ttk.Style()
style.theme_use("clam")

style.configure("Treeview", rowheight=28, font=("Segoe UI", 10))
style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))

# ---------------- LOAD DATA ----------------
pending_df = load_data(PENDING_FILE)
executed_df = load_data(EXECUTED_FILE)

# ---------------- ORDER ENTRY FRAME ----------------
entry_frame = ttk.LabelFrame(root, text="Order Entry", padding=15)
entry_frame.pack(fill="x", padx=20, pady=10)

# Party
ttk.Label(entry_frame, text="Party").grid(row=0, column=0, padx=10, pady=5)
party_cb = ttk.Combobox(entry_frame, values=PARTIES, state="readonly", width=20)
party_cb.grid(row=0, column=1)

# Item
ttk.Label(entry_frame, text="Item").grid(row=0, column=2, padx=10)
item_cb = ttk.Combobox(entry_frame, values=list(ITEMS.keys()), state="readonly", width=20)
item_cb.grid(row=0, column=3)

# Quantity
ttk.Label(entry_frame, text="Qty").grid(row=0, column=4, padx=10)
qty_entry = ttk.Entry(entry_frame, width=10)
qty_entry.grid(row=0, column=5)

# Price
ttk.Label(entry_frame, text="Price").grid(row=0, column=6, padx=10)
price_entry = ttk.Entry(entry_frame, width=10)
price_entry.grid(row=0, column=7)

def auto_price(event):
    item = item_cb.get()
    if item in ITEMS:
        price_entry.delete(0, tk.END)
        price_entry.insert(0, ITEMS[item])

item_cb.bind("<<ComboboxSelected>>", auto_price)

# ---------------- SAVE ORDER ----------------
def save_order():
    global pending_df

    try:
        party = party_cb.get()
        item = item_cb.get()
        qty = int(qty_entry.get())
        price = float(price_entry.get())

        if not party or not item:
            raise ValueError

        total = qty * price
        now = datetime.now()

        order_id = f"ORD{len(pending_df) + len(executed_df) + 1}"

        new_row = {
            "Order Id": order_id,
            "Party": party,
            "Item": item,
            "Qty": qty,
            "Price": price,
            "Total": total,
            "Order Date": now.date(),
            "Order Time": now.strftime("%H:%M:%S")
        }

        pending_df = pd.concat([pending_df, pd.DataFrame([new_row])], ignore_index=True)
        save_data(pending_df, PENDING_FILE)
        refresh_tables()

        qty_entry.delete(0, tk.END)
        messagebox.showinfo("Success", "Order saved as Pending")

    except:
        messagebox.showerror("Error", "Please enter valid order details")

ttk.Button(entry_frame, text="Save Order", command=save_order).grid(
    row=0, column=8, padx=20
)

# ---------------- TABLE AREA ----------------
table_frame = ttk.Frame(root)
table_frame.pack(fill="both", expand=True, padx=20, pady=10)

# ---------------- PENDING ----------------
pending_box = ttk.LabelFrame(table_frame, text="Pending Orders", padding=10)
pending_box.pack(side="left", fill="both", expand=True, padx=5)

pending_tree = ttk.Treeview(pending_box, columns=COLUMNS, show="headings")
for col in COLUMNS:
    pending_tree.heading(col, text=col)
    pending_tree.column(col, anchor="center", width=110)
pending_tree.pack(fill="both", expand=True)

# ---------------- EXECUTED ----------------
executed_box = ttk.LabelFrame(table_frame, text="Executed Orders", padding=10)
executed_box.pack(side="right", fill="both", expand=True, padx=5)

executed_tree = ttk.Treeview(executed_box, columns=COLUMNS, show="headings")
for col in COLUMNS:
    executed_tree.heading(col, text=col)
    executed_tree.column(col, anchor="center", width=110)
executed_tree.pack(fill="both", expand=True)

# ---------------- REFRESH ----------------
def refresh_tables():
    pending_tree.delete(*pending_tree.get_children())
    executed_tree.delete(*executed_tree.get_children())

    for _, row in pending_df.iterrows():
        pending_tree.insert("", "end", values=list(row))

    for _, row in executed_df.iterrows():
        executed_tree.insert("", "end", values=list(row))

refresh_tables()

# ---------------- EXECUTE ORDER ----------------
def execute_order():
    global pending_df, executed_df

    selected = pending_tree.selection()
    if not selected:
        messagebox.showwarning("Select Order", "Please select a pending order")
        return

    values = pending_tree.item(selected[0])["values"]
    executed_df = pd.concat([executed_df, pd.DataFrame([dict(zip(COLUMNS, values))])])
    pending_df = pending_df[pending_df["Order Id"] != values[0]]

    save_data(pending_df, PENDING_FILE)
    save_data(executed_df, EXECUTED_FILE)

    refresh_tables()
    messagebox.showinfo("Success", "Order executed successfully")

ttk.Button(
    pending_box,
    text="Execute Selected Order",
    command=execute_order
).pack(pady=10)

# ---------------- RUN ----------------
root.mainloop()
