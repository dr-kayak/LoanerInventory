from tkinter import simpledialog
import pyodbc
import tkinter as tk
from tkinter import messagebox, simpledialog, Toplevel, Label
from tkinter import ttk
from tkinter import Menu
from tkinter.simpledialog import askstring
import socket
import getpass

program_version = "1.1.0"


# Database connection setup
db_path = r'\\srvfileshare\Departments\9360 - Information Technology\Databases\LoanerInventory\LoanDatabase.accdb'
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=' + db_path + ';'
)

try:
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
except pyodbc.Error as e:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Database Error", f"Failed to connect to the database: {e}")
    exit()


def fetch_hardware_status():
    cursor.execute("SELECT User, Hardware, LoanDate, ReturnDate, Department, Phone, DeviceType, CrtUser FROM HardwareLoans")
    return cursor.fetchall()


def fetch_all_inventory():
    cursor.execute("SELECT Hardware, DeviceType, OnLoan, Serial, InitialDeployment FROM HardwareInventory")
    return cursor.fetchall()

def add_loan():
    user = user_entry.get()
    hardware = hardware_entry.get()
    loan_date = loan_date_entry.get()
    return_date = return_date_entry.get()
    department = department_entry.get()
    phone = phone_entry.get()
    device_type = device_type_entry.get()

    if user and hardware and loan_date and return_date and department and phone and device_type:
        try:
            cursor.execute(
                "INSERT INTO HardwareLoans (User, Hardware, LoanDate, ReturnDate, Department, Phone, DeviceType) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (user, hardware, loan_date, return_date, department, phone, device_type)
            )
            conn.commit()
            messagebox.showinfo("Success", "Loan record added successfully!")
            clear_entries()
            update_hardware_status()
            update_all_inventory()
        except pyodbc.Error as e:
            messagebox.showerror("Database Error", f"Failed to add record: {e}")
    else:
        messagebox.showwarning("Input Error", "Please fill in all fields")

def clear_entries():
    user_entry.delete(0, tk.END)
    hardware_entry.delete(0, tk.END)
    loan_date_entry.delete(0, tk.END)
    return_date_entry.delete(0, tk.END)
    department_entry.delete(0, tk.END)
    phone_entry.delete(0, tk.END)
    device_type_entry.delete(0, tk.END)

def open_add_entry_window():
    add_window = tk.Toplevel(root)
    add_window.title("Add Loan Record")
    add_window.geometry("400x400")
    add_window.resizable(True, True)  # Make the dialog box resizable

    frame = ttk.Frame(add_window, padding="10 10 10 10")
    frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    ttk.Label(frame, text="User:").grid(row=0, column=0, sticky=tk.W, pady=5)
    global user_entry
    user_entry = ttk.Entry(frame, width=30)
    user_entry.grid(row=0, column=1, pady=5)

    ttk.Label(frame, text="Hardware:").grid(row=1, column=0, sticky=tk.W, pady=5)
    global hardware_entry
    hardware_entry = ttk.Entry(frame, width=30)
    hardware_entry.grid(row=1, column=1, pady=5)

    ttk.Label(frame, text="Loan Date (YYYY-MM-DD):").grid(row=2, column=0, sticky=tk.W, pady=5)
    global loan_date_entry
    loan_date_entry = ttk.Entry(frame, width=30)
    loan_date_entry.grid(row=2, column=1, pady=5)

    ttk.Label(frame, text="Return Date (YYYY-MM-DD):").grid(row=3, column=0, sticky=tk.W, pady=5)
    global return_date_entry
    return_date_entry = ttk.Entry(frame, width=30)
    return_date_entry.grid(row=3, column=1, pady=5)

    ttk.Label(frame, text="Department:").grid(row=4, column=0, sticky=tk.W, pady=5)
    global department_entry
    department_entry = ttk.Entry(frame, width=30)
    department_entry.grid(row=4, column=1, pady=5)

    ttk.Label(frame, text="Phone:").grid(row=5, column=0, sticky=tk.W, pady=5)
    global phone_entry
    phone_entry = ttk.Entry(frame, width=30)
    phone_entry.grid(row=5, column=1, pady=5)

    ttk.Label(frame, text="Device Type:").grid(row=6, column=0, sticky=tk.W, pady=5)
    global device_type_entry
    device_type_entry = ttk.Entry(frame, width=30)
    device_type_entry.grid(row=6, column=1, pady=5)

    add_button = ttk.Button(frame, text="Add Loan Record", command=add_loan)
    add_button.grid(row=7, column=0, columnspan=2, pady=10)

def update_hardware_status():
    for i in loans_tree.get_children():
        loans_tree.delete(i)
    hardware_status = fetch_hardware_status()
    for row in hardware_status:
        if len(row) < 8:
            print(f"Error: Row doesn't have enough elements\nRow: {row}")
            continue
        formatted_loan_date = row[2].strftime('%Y-%m-%d') if row[2] else ''
        formatted_return_date = row[3].strftime('%Y-%m-%d') if row[3] else ''
        loans_tree.insert('', tk.END, values=(row[0], row[1], formatted_loan_date, formatted_return_date, row[4], row[5], row[6], row[7]))

loans_columns = ("User", "Hardware", "Loan Date", "Return Date", "Department", "Phone", "Device Type", "Loaning Tech")





def update_all_inventory():
    for i in inventory_tree.get_children():
        inventory_tree.delete(i)
    all_inventory = fetch_all_inventory()
    for row in all_inventory:
        formatted_initial_deployment = row[4].strftime('%Y-%m-%d') if row[4] else ''
        inventory_tree.insert('', tk.END, values=(row[0], row[1], row[2], row[3], formatted_initial_deployment))

from tkinter.simpledialog import askstring

from datetime import datetime, timedelta

import getpass

def copy_to_loans():
    selected_item = inventory_tree.selection()
    if selected_item:
        values = inventory_tree.item(selected_item, 'values')
        hardware, device_type, on_loan, serial, initial_deployment = values
        if on_loan == "No":  # Check if item is not already on loan
            # Create a Toplevel window for input
            copy_window = tk.Toplevel(root)
            copy_window.title("Loan Device")

            # Default loan date as current date
            current_date = datetime.now().strftime("%m/%d/%Y")
            
            # Labels and Entry fields for user input
            tk.Label(copy_window, text="User:").grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
            user_entry = tk.Entry(copy_window)
            user_entry.grid(row=0, column=1, padx=10, pady=5, sticky=tk.W)

            tk.Label(copy_window, text="Loan Date (MM/DD/YYYY):").grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)
            loan_date_entry = tk.Entry(copy_window)
            loan_date_entry.insert(0, current_date)  # Default loan date
            loan_date_entry.grid(row=1, column=1, padx=10, pady=5, sticky=tk.W)

            tk.Label(copy_window, text="Return Date (MM/DD/YYYY):").grid(row=2, column=0, padx=10, pady=5, sticky=tk.W)
            return_date_entry = tk.Entry(copy_window)
            return_date_entry.grid(row=2, column=1, padx=10, pady=5, sticky=tk.W)

            def set_default_period():
                loan_date_str = loan_date_entry.get()
                try:
                    loan_date = datetime.strptime(loan_date_str, "%m/%d/%Y")
                    return_date = loan_date + timedelta(days=30)
                    return_date_entry.delete(0, tk.END)
                    return_date_entry.insert(0, return_date.strftime("%m/%d/%Y"))
                except ValueError:
                    messagebox.showerror("Invalid Date", "Please enter a valid loan date (MM/DD/YYYY).")

            default_period_button = ttk.Button(copy_window, text="Default Period", command=set_default_period)
            default_period_button.grid(row=2, column=2, padx=10, pady=5, sticky=tk.W)

            tk.Label(copy_window, text="Department:").grid(row=3, column=0, padx=10, pady=5, sticky=tk.W)
            department_entry = tk.Entry(copy_window)
            department_entry.grid(row=3, column=1, padx=10, pady=5, sticky=tk.W)

            tk.Label(copy_window, text="Phone:").grid(row=4, column=0, padx=10, pady=5, sticky=tk.W)
            phone_entry = tk.Entry(copy_window)
            phone_entry.grid(row=4, column=1, padx=10, pady=5, sticky=tk.W)

            def submit_copy():
                user = user_entry.get()
                loan_date = loan_date_entry.get()
                return_date = return_date_entry.get()
                department = department_entry.get()
                phone = phone_entry.get()

                if user and loan_date and return_date and department and phone:
                    try:
                        cursor.execute(
                            "UPDATE HardwareInventory SET OnLoan = ? WHERE Hardware = ?",
                            ("Yes", hardware)
                        )
                        conn.commit()

                        # Retrieve the username of the currently logged-in user
                        current_user = getpass.getuser()

                        cursor.execute(
                            "INSERT INTO HardwareLoans (User, Hardware, LoanDate, ReturnDate, Department, Phone, DeviceType, CrtUser) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                            (user, hardware, loan_date, return_date, department, phone, device_type, current_user)
                        )
                        conn.commit()

                        update_all_inventory()
                        update_hardware_status()

                        messagebox.showinfo("Success", f"Item '{hardware}' added to Current Loans.")
                        copy_window.destroy()  # Close the dialog box after successful copy
                    except pyodbc.Error as e:
                        messagebox.showerror("Database Error", f"Failed to update record: {e}")
                else:
                    messagebox.showwarning("Incomplete Input", "Please fill in all fields.")

            submit_button = ttk.Button(copy_window, text="Submit", command=submit_copy)
            submit_button.grid(row=5, columnspan=2, padx=10, pady=10)
        else:
            messagebox.showwarning("Already on Loan", "This item is already on loan.")
    else:
        messagebox.showwarning("No Selection", "Please select an item to copy to Current Loans.")








def remove_loan():
    selected_item = loans_tree.selection()
    if selected_item:
        values = loans_tree.item(selected_item, 'values')
        hardware = values[1]  # Get the hardware name
        try:
            # Update "OnLoan" field to "No" in the "HardwareInventory" table
            cursor.execute(
                "UPDATE HardwareInventory SET OnLoan = ? WHERE Hardware = ?",
                ("No", hardware)
            )
            conn.commit()

            # Delete the corresponding entry from the "HardwareLoans" table
            cursor.execute(
                "DELETE FROM HardwareLoans WHERE Hardware = ?",
                (hardware,)
            )
            conn.commit()

            # Update the displays
            update_all_inventory()
            update_hardware_status()

            messagebox.showinfo("Success", f"Item '{hardware}' removed from Current Loans.")
        except pyodbc.Error as e:
            messagebox.showerror("Database Error", f"Failed to remove record: {e}")
    else:
        messagebox.showwarning("No Selection", "Please select an item to remove from Current Loans.")

def open_add_item_window():
    add_window = tk.Toplevel(root)
    add_window.title("Add New Item")
    add_window.geometry("400x300")
    add_window.resizable(True, True)  # Make the dialog box resizable

    frame = ttk.Frame(add_window, padding="10 10 10 10")
    frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    # Label and Entry for Hardware
    ttk.Label(frame, text="Hardware:").grid(row=0, column=0, sticky=tk.W, pady=5)
    hardware_entry = ttk.Entry(frame, width=30)
    hardware_entry.grid(row=0, column=1, pady=5)

    # Label and Entry for Device Type
    ttk.Label(frame, text="Device Type:").grid(row=1, column=0, sticky=tk.W, pady=5)
    device_type_entry = ttk.Entry(frame, width=30)
    device_type_entry.grid(row=1, column=1, pady=5)

    # Label and Entry for Serial
    ttk.Label(frame, text="Serial:").grid(row=2, column=0, sticky=tk.W, pady=5)
    serial_entry = ttk.Entry(frame, width=30)
    serial_entry.grid(row=2, column=1, pady=5)

    # Label and Entry for Initial Deployment
    ttk.Label(frame, text="Initial Deployment (YYYY-MM-DD):").grid(row=3, column=0, sticky=tk.W, pady=5)
    initial_deployment_entry = ttk.Entry(frame, width=30)
    initial_deployment_entry.grid(row=3, column=1, pady=5)

    # Set default value for OnLoan field
    on_loan_value = "No"

    def add_new_item():
        hardware = hardware_entry.get()
        device_type = device_type_entry.get()
        serial = serial_entry.get()
        initial_deployment = initial_deployment_entry.get()

        if hardware and device_type and serial and initial_deployment:
            try:
                cursor.execute(
                    "INSERT INTO HardwareInventory (Hardware, DeviceType, OnLoan, Serial, InitialDeployment) VALUES (?, ?, ?, ?, ?)",
                    (hardware, device_type, on_loan_value, serial, initial_deployment)
                )
                conn.commit()
                messagebox.showinfo("Success", "New item added successfully!")
                add_window.destroy()  # Close the dialog window after successful addition
                update_all_inventory()  # Update the display
            except pyodbc.Error as e:
                messagebox.showerror("Database Error", f"Failed to add item: {e}")
        else:
            messagebox.showwarning("Input Error", "Please fill in all fields")

    # Button to add the new item
    add_button = ttk.Button(frame, text="Add New Item", command=add_new_item)
    add_button.grid(row=4, column=0, columnspan=2, pady=10)



# Update the display after adding a new item
    update_all_inventory()

def get_about_text():
    return (
        f"Loaner Inventory System v{program_version}\n"
        "Developed by Malachi McRee\n\n"
        "Version History:\n"
        "Loaner Inventory System\n"
        "Version 1.1.0\n"
        "Changelog:\n"
        "- Removed CMD from showing when program was running\n"
        "- Added Changelog\n"
        "- Now I'll know if people ACTUALLY have the latest version\n"
        "- Moved the database so Kendal will chill\n"
        "- Might start naming versions after stupid nature stuff like Apple\n\n"
        
        "Loaner Inventory System\n"
        "Version 1.0.0\n"
        "Changelog:\n"
        "- Inital Release (6/12/24)\n"
        "- Barebones, but functional. Sue me.\n"
    )

def show_about():
    about_window = Toplevel(root)
    about_window.title("About Loaner Inventory System")
    about_window.geometry("500x550")
    
    about_label = Label(about_window, text=get_about_text(), justify="left")
    about_label.pack(pady=20, padx=20)

root = tk.Tk()
root.title(f"IT Hardware Loan Tracker - Version {program_version}")
root.geometry("1000x500")
root.resizable(True, True)



# Create a style object for the dark mode theme
dark_style = ttk.Style()
dark_style.theme_use('clam')  # Use 'clam' theme as base

# Configure the colors for dark mode
dark_style.configure('TFrame', background='#2e2e2e')
dark_style.configure('TLabel', background='#2e2e2e', foreground='#ffffff')
dark_style.configure('TEntry', fieldbackground='#454545', foreground='#ffffff')
dark_style.configure('TButton', background='#454545', foreground='#ffffff')
dark_style.configure('Treeview', background='#2e2e2e', foreground='#ffffff', fieldbackground='#454545')
dark_style.configure('Treeview.Heading', background='#454545', foreground='#ffffff')

# Notebook for tabs
notebook = ttk.Notebook(root)
notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Frame for Current Loans tab
loans_frame = ttk.Frame(notebook, padding="10 10 10 10")
loans_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Treeview for hardware status in Current Loans tab
loans_columns = ("User", "Hardware", "Loan Date", "Return Date", "Department", "Phone", "Device Type", "Loaning Tech")
loans_tree = ttk.Treeview(loans_frame, columns=loans_columns, show='headings')
for col in loans_columns:
    loans_tree.heading(col, text=col)
    loans_tree.column(col, width=120)
loans_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Refresh Button in Current Loans tab
refresh_loans_button = ttk.Button(loans_frame, text="Refresh", command=update_hardware_status)
refresh_loans_button.grid(row=3, column=0, pady=10)

# Remove Loan Button in Current Loans tab
remove_loan_button = ttk.Button(loans_frame, text="Remove Loan", command=remove_loan)
remove_loan_button.grid(row=2, column=0, pady=10)

# Frame for All Inventory tab
inventory_frame = ttk.Frame(notebook, padding="10 10 10 10")

# Treeview for all inventory in All Inventory tab
inventory_columns = ("Hardware", "Device Type", "On Loan", "Serial", "Initial Deployment")
inventory_tree = ttk.Treeview(inventory_frame, columns=inventory_columns, show='headings')
for col in inventory_columns:
    inventory_tree.heading(col, text=col)
    inventory_tree.column(col, width=140)
inventory_tree.grid(row=0, column=0, padx=150, pady=10, sticky=(tk.W, tk.E, tk.N, tk.S))

# Refresh Button in All Inventory tab
refresh_inventory_button = ttk.Button(inventory_frame, text="Refresh", command=update_all_inventory)
refresh_inventory_button.grid(row=1, column=0, pady=10)

# Add New Item Button in All Inventory tab
add_item_button = ttk.Button(inventory_frame, text="Add New Item", command=open_add_item_window)
add_item_button.grid(row=2, column=0, pady=10)

# Copy to Loans Button in All Inventory tab
copy_to_loans_button = ttk.Button(inventory_frame, text="Loan Device", command=copy_to_loans)
copy_to_loans_button.grid(row=3, column=0, pady=10)

# Create the About button
about_button = tk.Button(root, text="About", command=show_about)
about_button.place(x=940, y=0)  # Adjust the coordinates as necessary

# Adding tabs to the notebook
notebook.add(loans_frame, text="Current Loans")
notebook.add(inventory_frame, text="All Inventory")

# Define the function to remove a device from the inventory
def remove_device():
    selected_item = inventory_tree.selection()
    if selected_item:
        values = inventory_tree.item(selected_item, 'values')
        serial = values[3]  # Get the serial number of the device

        # Create a Toplevel window for input
        remove_window = tk.Toplevel(root)
        remove_window.title("Remove Device")

        tk.Label(remove_window, text="Issued User:").grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
        issued_user_entry = tk.Entry(remove_window)
        issued_user_entry.grid(row=0, column=1, padx=10, pady=5, sticky=tk.W)

        tk.Label(remove_window, text="New Name:").grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)
        new_name_entry = tk.Entry(remove_window)
        new_name_entry.grid(row=1, column=1, padx=10, pady=5, sticky=tk.W)

        def submit_removal():
            issued_user = issued_user_entry.get()
            new_name = new_name_entry.get()

            # Get the current user
            current_user = getpass.getuser()

            if issued_user and new_name:
                try:
                    # Insert the data into the RemovedDevices table
                    cursor.execute(
                        "INSERT INTO RemovedDevices (Serial, CrtUser, User, NewName) VALUES (?, ?, ?, ?)",
                        (serial, current_user, issued_user, new_name)
                    )
                    conn.commit()

                    # Delete the device from the HardwareInventory table
                    cursor.execute(
                        "DELETE FROM HardwareInventory WHERE Serial = ?",
                        (serial,)
                    )
                    conn.commit()

                    # Update the display
                    update_all_inventory()
                    messagebox.showinfo("Success", f"Device with serial '{serial}' removed successfully.")
                    remove_window.destroy()  # Close the dialog box after successful removal
                except pyodbc.Error as e:
                    messagebox.showerror("Database Error", f"Failed to remove device: {e}")
            else:
                messagebox.showwarning("Incomplete Input", "Please fill in all fields.")

        submit_button = ttk.Button(remove_window, text="Submit", command=submit_removal)
        submit_button.grid(row=2, columnspan=2, padx=10, pady=10)
    else:
        messagebox.showwarning("No Selection", "Please select a device to remove.")





# Add the "Remove Device" button under the All Inventory tab
remove_device_button = ttk.Button(inventory_frame, text="Remove Device", command=remove_device)
remove_device_button.grid(row=4, column=0, pady=10)





root.mainloop()

# Close the database connection when done
conn.close()