"""EmailDistroBuilder - Copyright 2021, William Jenkins
Program is written to extract email addresses from an Excel (.xlsx) file. The
program is written specific to the file's format.  Email address columns must
be adjusted if the standard format is not used.
"""

import tkinter as tk
from tkinter import filedialog, ttk

import pandas as pd


class Emails:
    """Stores attributes for file containing email addresses."""
    def __init__(self):
        self.filename = None
        self.option = []
        self.address_list = []


def browseFiles():
    """Opens file browser and sets file path."""
    filename = filedialog.askopenfilename(
        initialdir = "/",
        title = "Select a File",
        filetypes = (("Excel Files", "*.xlsx"), ("all files", "*.*"))
    )
    try:
        label_file_explorer.configure(text="File Opened: "+filename)
        Emails.filename = filename
    except:
        return


def emailsToClipboard():
    """Copies email addresses to system clipboard."""
    window.clipboard_clear()
    window.clipboard_append(Emails.address_list)


def emailsToFile():
    """Saves email addresses to user-specified file."""
    f = filedialog.asksaveasfile(mode="w", defaultextension=".txt")
    if f is None:
        return
    f.write(Emails.address_list)
    f.close()


def getAddresses():
    """Extracts email addresses from Excel file."""
    try:
        df = pd.read_excel(Emails.filename)
        option = clicked.get()

        if option == "Military":
            column = df.iloc[:,12]
            addresses = [item for item in column if ("@" in str(item) and ".mil" in str(item))]
        elif option == "Civilian":
            column = df.iloc[:,13]
            addresses = [item for item in column if ("@" in str(item) and ".mil" not in str(item))]
        elif option == "All":
            column = pd.concat([df.iloc[:,12], df.iloc[:,13]])
            addresses = [item for item in column if "@" in str(item)]

        Emails.address_list = ';'.join(addresses)
        extract_status.configure(text="Status: "+"\U00002705")
    except:
        if not hasattr(Emails, "filename"):
            errmsg = "File not selected."
        else:
            errmsg = "Unable to extract email addresses."
        tk.messagebox.showerror(title="Error!", message=errmsg)
        extract_status.configure(text="Status: "+"\U0000274C")


if __name__ == "__main__":
    window = tk.Tk()
    window.title("Email Distro Builder")
    window.geometry("500x300")
    window.columnconfigure(0, minsize=500)
    # canvas = tk.Canvas(window, height=300, width=500)

    # File Explorer ===========================================================
    loadpath_label = tk.Label(
        text="Select location of recall list:",
        fg="black"
    ).grid(row=0)
    # loadpath_label.pack()

    loadpath_button = tk.Button(
        window,
        text="Browse Files",
        command=browseFiles,
        fg="black"
    ).grid(row=1)
    # loadpath_button.pack()

    label_file_explorer = tk.Label(
        window,
        text="File Opened: ",
        fg="black"
    )
    label_file_explorer.grid(row=2)
    # label_file_explorer.pack()

    # Dropdown Menu for Email Address Type ====================================
    line1 = ttk.Separator(
        window,
        orient="horizontal"
    ).grid(column=0, row=3, sticky="ew")

    label_options = tk.Label(
        window,
        text="Select Military/Civilian/All email addresses:",
        fg="black"
    ).grid(row=4)
    # label_options.pack()

    options = ["Military", "Civilian", "All"]

    clicked = tk.StringVar()
    clicked.set(options[0])
    menu = tk.OptionMenu(window, clicked, *options)
    menu.config(fg="black")
    menu.grid(row=5)
    # menu.pack()

    # # Get Emails ==============================================================
    extract_button = tk.Button(
        window,
        text="Extract Email Addresses",
        command=getAddresses,
        fg="black"
    ).grid(row=6)
    # extract_button.pack()

    extract_status = tk.Label(
        window,
        text="Status: ",
        fg="black"
    )
    extract_status.grid(row=7)

    # Copy to Clipboard =======================================================
    line2 = ttk.Separator(
        window,
        orient="horizontal"
    ).grid(column=0, row=8, sticky="ew")

    copy_button = tk.Button(
        window,
        text="Copy Addresses to Clipboard",
        command=emailsToClipboard,
        fg="black"
    ).grid(row=9)
    # copy_button.pack()

    # Save to File ============================================================
    save_button = tk.Button(
        window,
        text="Save as Text File",
        command=emailsToFile,
        fg="black"
    ).grid(row=10)
    # save_button.pack()

    # Credit ==================================================================
    line3 = ttk.Separator(
        window,
        orient="horizontal"
    ).grid(column=0, row=11, sticky="ew")

    credit_label = tk.Label(
        window,
        text="\U000000A9"+"2021 William Jenkins",
        fg="black"
    ).grid(row=12)

    # Initiate Window Loop ====================================================
    window.mainloop()
