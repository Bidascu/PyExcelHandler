from tkinter import *
from tkinter import filedialog
import pandas as pd

# Global filename -- In order to use variable filename in multiple functions, it is declared as global 
filename = ""

# Function to swap sheet info
def swap():
    global filename
    if not filename:
        label_file_explorer.configure(text="No file selected. Please choose a file first.")
        return

    try:
        # Read the sheets
        df1 = pd.read_excel(io=filename, sheet_name='Sheet1')
        df2 = pd.read_excel(io=filename, sheet_name='Sheet2')

        # Swap the sheets
        with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists="replace") as writer:
            df1.to_excel(writer, sheet_name='Sheet2', index=False)
            df2.to_excel(writer, sheet_name='Sheet1', index=False)

        label_file_explorer.configure(text="Sheets swapped successfully!")

    except Exception as e:
        label_file_explorer.configure(text=f"Error: {str(e)}")

# Function for file selection
def browseFiles():
    global filename
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select a File",
                                          filetypes=(("Excel files", "*.xlsx*"),
                                                     ("All files", "*.*")))
    if filename:
        label_file_explorer.configure(text="File Opened: " + filename)
    else:
        label_file_explorer.configure(text="No file selected.")

# Create the root window
window = Tk()
window.title('Sheet Swapper')
window.geometry("700x200")
window.config(background="#FFFFFF")

# UI Elements
label_file_explorer = Label(window, text="Select a File (Using Tkinter)", width=100, height=4, fg="blue")
button_explore = Button(window, text="Browse Files", command=browseFiles)
button_exit = Button(window, text="Exit", command=window.quit)
button_swap = Button(window, text="Swap", command=swap)

# Layout
label_file_explorer.grid(column=1, row=1)
button_explore.grid(column=1, row=2)
button_exit.grid(column=1, row=3)
button_swap.grid(column=1, row=4)

# Run GUI
window.mainloop()