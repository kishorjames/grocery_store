import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook
from reportlab.pdfgen import canvas

# Function to calculate total cost and save details
def save_details():
    name = name_var.get()
    contact = contact_var.get()
    address = address_var.get()
    selected_items = []
    total_cost = 0

    # Add selected items and their prices
    for item, var in item_vars.items():
        if var.get() == 1:
            selected_items.append(item)
            total_cost += prices[item]

    if not name or not contact or not address or not selected_items:
        messagebox.showerror("Error", "Please fill in all details and select items!")
        return

    # Save to Excel
    data = [name, contact, ", ".join(selected_items), total_cost, address]
    ws.append(data)
    wb.save("Grocery_Details.xlsx")

    # Save Receipt as PDF
    save_receipt_as_pdf(name, contact, selected_items, total_cost, address)

    messagebox.showinfo("Success", "Order placed successfully!")
    reset_form()

# Function to reset form
def reset_form():
    name_var.set("")
    contact_var.set("")
    address_var.set("")
    for var in item_vars.values():
        var.set(0)

# Function to save receipt as PDF
def save_receipt_as_pdf(name, contact, items, total_cost, address):
    receipt_filename = f"Receipt_{name}.pdf"
    c = canvas.Canvas(receipt_filename)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, 800, "Grocery Store Receipt")
    c.setFont("Helvetica", 12)
    c.drawString(50, 770, f"Customer Name: {name}")
    c.drawString(50, 750, f"Contact Number: {contact}")
    c.drawString(50, 730, f"Address: {address}")
    c.drawString(50, 710, f"Items Purchased:")
    y = 690
    for item in items:
        c.drawString(70, y, f"- {item} ({prices[item]} Rs)")
        y -= 20
    c.drawString(50, y - 20, f"Total Cost: {total_cost} Rs")
    c.save()

# Initialize tkinter window
root = tk.Tk()
root.title("Grocery Selection Form")
root.geometry("600x600")

# Excel Setup
wb = Workbook()
ws = wb.active
ws.append(["Customer Name", "Contact No.", "Items Purchased", "Total Cost", "Address"])

# Prices of items
prices = {
    "Rice": 50,
    "Grains": 40,
    "Millets": 60,
    "Milk": 25,
    "Eggs": 6,
    "Canned Beans": 45,
    "Tomato Ketchup": 70,
    "Soya Sauce": 85,
    "Honey": 150
}

# Variables for form inputs
name_var = tk.StringVar()
contact_var = tk.StringVar()
address_var = tk.StringVar()
item_vars = {item: tk.IntVar() for item in prices.keys()}

# GUI Design
tk.Label(root, text="Customer Name").pack(anchor="w", padx=20, pady=5)
tk.Entry(root, textvariable=name_var, width=40).pack(anchor="w", padx=20)

tk.Label(root, text="Contact Number").pack(anchor="w", padx=20, pady=5)
tk.Entry(root, textvariable=contact_var, width=40).pack(anchor="w", padx=20)

tk.Label(root, text="Address").pack(anchor="w", padx=20, pady=5)
tk.Entry(root, textvariable=address_var, width=40).pack(anchor="w", padx=20)

tk.Label(root, text="Select Grocery Items:").pack(anchor="w", padx=20, pady=5)
for item, var in item_vars.items():
    tk.Checkbutton(root, text=f"{item} ({prices[item]} Rs)", variable=var).pack(anchor="w", padx=40)

tk.Button(root, text="Place Order", command=save_details, bg="green", fg="white").pack(pady=20)
tk.Button(root, text="Reset Form", command=reset_form, bg="gray", fg="white").pack()

root.mainloop()
