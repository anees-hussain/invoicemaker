import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import pandas as pd
from fpdf import FPDF
import os
import sqlite3
import win32print
import tempfile

# Initialize database and create products table if it doesn't exist
def initialize_db():
    conn = sqlite3.connect("products.db")
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS products (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        rate INTEGER NOT NULL
    )
    """)
    conn.commit()
    conn.close()

# Function to get all products from the database
def get_all_products():
    conn = sqlite3.connect("products.db")
    cursor = conn.cursor()
    cursor.execute("SELECT name, rate FROM products")
    products = cursor.fetchall()
    conn.close()
    return products

# Function to add a new product to the database
def add_product_to_db(name, rate):
    conn = sqlite3.connect("products.db")
    cursor = conn.cursor()
    cursor.execute("INSERT INTO products (name, rate) VALUES (?, ?)", (name, rate))
    conn.commit()
    conn.close()

# Function to delete a product from the database
def delete_product_from_db(name):
    conn = sqlite3.connect("products.db")
    cursor = conn.cursor()
    cursor.execute("DELETE FROM products WHERE name=?", (name,))
    conn.commit()
    conn.close()

# Function to update a product in the database
def update_product_in_db(name, new_rate):
    conn = sqlite3.connect("products.db")
    cursor = conn.cursor()
    cursor.execute("UPDATE products SET rate=? WHERE name=?", (new_rate, name))
    conn.commit()
    conn.close()

# Function to upload and process the Excel file
def upload_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")],
        title="Select an Excel File"
    )
    
    if file_path:
        loading_label.config(text="Loading... Please wait.")
        root.update_idletasks()

        try:
            df = pd.read_excel(file_path)
            
            if df.empty:
                raise ValueError("The selected Excel file is empty.")
            
            required_columns = ['Customer Name', 'Address', 'Delivery Date'] + [product[0] for product in get_all_products()]
            if not all(col in df.columns for col in required_columns):
                missing_cols = [col for col in required_columns if col not in df.columns]
                raise ValueError(f"Excel file is missing the following columns: {', '.join(missing_cols)}")
            
            invoices = []

            # Iterate over each row to generate invoices
            for index, row in df.iterrows():
                customer_name = row['Customer Name']
                customer_address = row['Address']
                delivery_date = row['Delivery Date'].strftime("%d-%m-%Y")  # Format date to show only the date
                
                # Calculate total for each product in the current row
                total_lines = []
                grand_total = 0
                
                for product_name, product_rate in get_all_products():
                    if product_name in row and pd.notna(row[product_name]) and row[product_name] > 0:
                        order_quantity = int(row[product_name])
                        total_amount = int(order_quantity * product_rate)  # Convert to integer
                        total_lines.append(f"{product_name}: {order_quantity} x {product_rate} = {total_amount}")
                        grand_total += total_amount
                
                grand_total = int(grand_total)  # Convert grand total to integer
                invoice_text = f"Delivery Date: {delivery_date}\n"
                invoice_text += "\n"
                invoice_text += f"Customer Name: {customer_name}\nAddress: {customer_address}\n"
                invoice_text += "\n"
                invoice_text += '\n'.join(total_lines)
                invoice_text += f"\n\nGrand Total: {grand_total}\n"
                #invoice_text += "-" * 40  # Dotted line

                invoices.append(invoice_text)
            
            show_all_invoices(invoices)
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process the file: {e}")
        
        finally:
            loading_label.config(text="")  # Hide loading label when done

# Function to display all invoices in one window with scrolling
def show_all_invoices(invoices):
    invoice_window = tk.Toplevel()
    invoice_window.title("All Invoices")
    invoice_window.geometry("550x600")  # Increase the width to fit buttons

    # Create a canvas with a scrollbar
    canvas = tk.Canvas(invoice_window)
    scrollbar = ttk.Scrollbar(invoice_window, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    # Add all invoices to the scrollable frame
    for invoice in invoices:
        invoice_label = tk.Label(scrollable_frame, text=invoice, justify=tk.LEFT, padx=20, pady=10)
        invoice_label.pack()

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Save and print buttons
    save_button = tk.Button(invoice_window, text="Save as PDF", command=lambda: save_all_as_pdf(invoices))
    save_button.pack(side=tk.LEFT, padx=10, pady=10)

    print_button = tk.Button(invoice_window, text="Print", command=lambda: select_printer_and_print(invoices))
    print_button.pack(side=tk.RIGHT, padx=10, pady=10)

# Function to save all invoices as a single PDF
def save_all_as_pdf(invoices):
    file_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
        title="Save All Invoices As",
        initialfile="All_Invoices.pdf"
    )
    
    if file_path:
        try:
            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()
            pdf.set_font("Arial", size=13)
            
            invoices_per_page = 2
            invoice_count = 0
            
            for invoice in invoices:
                # Split invoice into lines and format them
                lines = invoice.strip().split('\n')
                for line in lines:
                    if any(label in line for label in ["Delivery Date", "Customer Name", "Address", "Grand Total"]):
                        if "Grand Total" in line:
                            pdf.set_font("Arial", "B", 12)  # Bold "Grand Total" heading only
                        else:
                            pdf.set_font("Arial", "B", 12)  # Bold headings
                            pdf.multi_cell(200, 10, txt=line.split(":")[0])
                            pdf.set_font("Arial", size=12)  # Regular text
                            pdf.cell(200, 10, txt=line.split(":")[1].strip(), ln=1)
                    else:
                        pdf.set_font("Arial", size=12)  # Regular text
                        pdf.cell(200, 10, txt=line, ln=True)

                invoice_count += 1
                if invoice_count == invoices_per_page:
                    invoice_count = 0
                    pdf.add_page()  # Start new page after every 2 invoices
                else:
                    pdf.cell(200, 10, txt="", ln=True)  # New line after each invoice
                    pdf.cell(200, 10, txt="-" * 100, ln=True)  # Dotted line

            pdf.output(file_path)
            messagebox.showinfo("Success", f"Invoices saved as {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save the PDF: {e}")

# Function to select a printer and print all invoices
def select_printer_and_print(invoices):
    printer_name = win32print.GetDefaultPrinter()
    printer_name = win32print.GetDefaultPrinter()

    if printer_name:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".txt") as tmp_file:
            for invoice in invoices:
                tmp_file.write(invoice.encode('utf-8'))
                tmp_file.write(b"\n" + b"-" * 83 + b"\n\n")  # Add separator

            tmp_file_path = tmp_file.name

        # Print the file
        os.startfile(tmp_file_path, "print")

# Function to manage the product database
def manage_products():
    def refresh_product_list():
        products = get_all_products()
        product_listbox.delete(0, tk.END)
        for product in products:
            product_listbox.insert(tk.END, f"{product[0]} - {product[1]}")

    def add_product():
        product_name = simpledialog.askstring("Input", "Enter the product name:")
        if product_name:
            try:
                product_rate = int(simpledialog.askstring("Input", f"Enter the rate for {product_name}:"))
                add_product_to_db(product_name.strip(), product_rate)
                refresh_product_list()
                messagebox.showinfo("Success", f"Product '{product_name}' added with rate {product_rate}")
            except ValueError:
                messagebox.showerror("Error", "Invalid rate entered. Please enter a numeric value.")

    def delete_product():
        selected_product = product_listbox.get(tk.ACTIVE)
        if selected_product:
            product_name = selected_product.split(" - ")[0]
            delete_product_from_db(product_name)
            refresh_product_list()
            messagebox.showinfo("Success", f"Product '{product_name}' deleted.")

    def update_product():
        selected_product = product_listbox.get(tk.ACTIVE)
        if selected_product:
            product_name = selected_product.split(" - ")[0]
            try:
                product_name = simpledialog.askstring("Input", "Enter the new name for product:")
                new_rate = int(simpledialog.askstring("Input", f"Enter the new rate for {product_name}:"))
                update_product_in_db(product_name.strip(), new_rate)
                refresh_product_list()
                messagebox.showinfo("Success", f"Rate for '{product_name}' updated to {new_rate}.")
            except ValueError:
                messagebox.showerror("Error", "Invalid rate entered. Please enter a numeric value.")

    manage_window = tk.Toplevel()
    manage_window.title("Manage Products")
    manage_window.geometry("400x400")

    product_listbox = tk.Listbox(manage_window, width=50, height=15)
    product_listbox.pack(pady=20)

    add_button = tk.Button(manage_window, text="Add Product", command=add_product)
    add_button.pack(pady=5)

    update_button = tk.Button(manage_window, text="Update Rate", command=update_product)
    update_button.pack(pady=5)

    delete_button = tk.Button(manage_window, text="Delete Product", command=delete_product)
    delete_button.pack(pady=5)

    refresh_product_list()

# Main application window
root = tk.Tk()
root.title("Shoaib Traders")
root.geometry("400x400")
root.configure(bg="white")

# Initialize the database
initialize_db()

# Loading label
loading_label = tk.Label(root, text="", fg="red")
loading_label.pack()

# Upload button
upload_button = tk.Button(root, text="Upload File", command=upload_file,
                          bg="green", fg="white", borderwidth=2, relief="solid")
upload_button.pack(pady=10)

# Manage Products button
manage_products_button = tk.Button(root, text="Manage Products", command=manage_products,
                                   bg="blue", fg="white", borderwidth=2, relief="solid")
manage_products_button.pack(pady=10)

root.mainloop()
