import tkinter as tk
from tkinter import filedialog, messagebox
import os
import openpyxl
from datetime import datetime

def select_file():
    file_path = filedialog.askopenfilename()
    renaming_window(file_path)

def sanitize_filename(filename):
    # Define a string containing all forbidden character
    forbidden_characters = r'\/:*?"<>,|'
    
    # Replace each forbidden character with a space
    for char in forbidden_characters:
        #print(char)
        filename = filename.replace(char, ' ')
    
    return filename

def look_up_project(project_number_var, order_number_var, client_name_var, proj_name_var, lookup_status_label, file_name):
    project_number = str(project_number_var.get())  # Ensure project number is string for comparison
    found = False
    try:
        workbook = openpyxl.load_workbook(r"C:\RT\Data\OPEN PROJECTS.xlsx")
        sheet = workbook.active
        for row_num, row in enumerate(sheet.iter_rows(values_only=True)):
            if str(row[1]) == project_number:
                order_number_var.set(row[0])  # Set order number
                client_name_var.set(row[3])  # Set client name
                proj_name_var.set(row[4])  # Set project name
                found = True
                break
        if not found:
            lookup_status_label.config(text="Project Number Not Found!", fg="red")
        else:
            lookup_status_label.config(text=f"Project Number Found in Row {row_num + 1}", fg="blue")

        # Schedule the text update to occur after 3 seconds (3000 milliseconds)
        lookup_status_label.after(3000, lambda: lookup_status_label.config(text=str(len(file_name.get())), fg="black"))
    except Exception as e:
        lookup_status_label.config(text=f"File Reading Error!", fg="red")
        lookup_status_label.after(3000, lambda: lookup_status_label.config(text=str(len(file_name.get())), fg="black"))

def select_file():
    file_path = filedialog.askopenfilename()
    renaming_window(file_path)

def rename_file(old_file_path, file_name_var, file_extension_var, rename_button, open_new_file_button):
    def sanitize_filename(filename):
        forbidden_characters = r'\/:*?"<>|'
        for char in forbidden_characters:
            filename = filename.replace(char, ' ')
        return filename

    file_name = sanitize_filename(file_name_var.get())
    file_extension = file_extension_var.get()  # assuming file_extension is also a StringVar object

    directory = os.path.dirname(old_file_path)
    new_file_path = os.path.join(directory, file_name + file_extension)

    try:
        os.rename(old_file_path, new_file_path)
    except FileNotFoundError as e:
        print(f"An error occurred: {e}")
        return

    rename_button.config(state=tk.DISABLED)
    open_new_file_button.config(state=tk.NORMAL)
    
    
def renaming_window(file_path):
    def update_filename(*args):
        file_name.set(
            f"{document_type_var.get()}_{project_number_var.get()}_{order_number_var.get()}_{client_name_var.get()}_{proj_name_var.get()}_{date_var.get()}_{revision_number_var.get()}"
        )
        lookup_status_label.config(text=str(len(file_name.get())), fg="black")



    window = tk.Tk()
    window.title("Red Thread File Renamer")

    # Window size
    width = 650
    height = 400

    # Calculate the position to center the window
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x_coordinate = (screen_width / 2) - (width / 2)
    y_coordinate = (screen_height / 2) - (height / 2)

    window.geometry(f"{width}x{height}+{int(x_coordinate)}+{int(y_coordinate)}")

    window.attributes('-topmost', True)  # Ensure window is on top

    arial_font = ("Arial", 10)
    button_font = ("Arial", 12)


    document_type_var = tk.StringVar()
    project_number_var = tk.StringVar()
    order_number_var = tk.StringVar()
    client_name_var = tk.StringVar()
    proj_name_var = tk.StringVar()
    date_var = tk.StringVar(value=datetime.now().strftime("%m-%d-%y"))
    revision_number_var = tk.StringVar(value="0")
    file_name = tk.StringVar()
    file_extension = tk.StringVar(value=os.path.splitext(file_path)[1])
    original_file_path = tk.StringVar(value=file_path)

    tk.Label(window, text="Document Type:", anchor=tk.W, font=arial_font).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
    with open(r"C:\RT\Data\document_types.txt", "r") as file:
        document_types = file.read().splitlines()
        document_type_var.set(document_types[0])
        document_type_var.trace("w", update_filename)
        option_menu = tk.OptionMenu(window, document_type_var, *document_types)
        option_menu.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        option_menu.config(width=25)

    # Project Number entry and Look Up button
    tk.Label(window, text="Project Num:", anchor=tk.W, font=arial_font).grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
    entry = tk.Entry(window, textvariable=project_number_var, font=arial_font)
    entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
    lookup_button = tk.Button(window, text="Look Up", command=lambda: look_up_project(project_number_var, order_number_var, client_name_var, proj_name_var, lookup_status_label, file_name), font=arial_font, width=14)
    lookup_button.grid(row=1, column=1, padx=5, pady=5)
    lookup_status_label = tk.Label(window, text="", font=arial_font, anchor='e')
    lookup_status_label.grid(row=2, column=1, padx=5, pady=5, sticky='e')
    project_number_var.trace("w", update_filename) # Add this line
    

    # Shift all the following rows down by one to make room for the Look Up button
    fields = [("Order Num:", order_number_var), ("Client Name:", client_name_var), ("Project Name:", proj_name_var), ("Date:", date_var), ("Revision:", revision_number_var)]
    for i, (label, var) in enumerate(fields):
        tk.Label(window, text=label, anchor=tk.W, font=arial_font).grid(row=i + 2, column=0, padx=5, pady=5, sticky=tk.W)
        entry = tk.Entry(window, textvariable=var, font=arial_font)
        entry.grid(row=i + 2, column=1, padx=5, pady=5, sticky=tk.W)
        var.trace("w", update_filename)
        

    update_filename()
    tk.Label(window, text="New File Name:", anchor=tk.W, font=arial_font).grid(row=len(fields) + 2, column=0, padx=5, pady=5, sticky=tk.W)
    entry_file_name = tk.Entry(window, textvariable=file_name, font=arial_font, width=60)
    entry_file_name.grid(row=len(fields) + 2, column=1, padx=5, pady=5, sticky=tk.W)
    tk.Label(window, text=" ", anchor=tk.W, font=arial_font).grid(row=len(fields) + 2, column=2, padx=0, pady=5, sticky=tk.W)
    tk.Entry(window, textvariable=file_extension, font=arial_font, width=5).grid(row=len(fields) + 2, column=3, padx=5, pady=5, sticky=tk.W)


    def open_new_file():
        window.destroy()
        select_file()
    
    # Create a frame for buttons
    button_frame = tk.Frame(window)
    button_frame.grid(row=len(fields) + 3, column=0, columnspan=4, pady=20)

    # Rename File button Renames the file after all field are entered, disables it self, and enables open new file 
    rename_button = tk.Button(button_frame, text="Rename", font=button_font, command=lambda: rename_file(original_file_path.get(), file_name, file_extension, rename_button, open_new_file_button), width=15)
    rename_button.grid(row=0, column=0, padx=10, pady=20)
    
    # Open New file button, only enabled after rename is clicked. opens new file to restart process. 
    open_new_file_button = tk.Button(button_frame, text="Open New File", font=button_font, command=open_new_file, state=tk.DISABLED, width=15)
    open_new_file_button.grid(row=0, column=1, padx=10, pady=20)
    
    # Done button to close the program
    done_button = tk.Button(button_frame, text="Done", font=button_font, command=window.destroy, width=15)
    done_button.grid(row=0, column=2, padx=10, pady=20)

    window.mainloop()

        
select_file()