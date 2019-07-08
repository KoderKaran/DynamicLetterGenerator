import win32print
import os, sys
import tkinter as tk
from tkinter import filedialog
import openpyxl as op
from PIL import ImageTk, Image
finishedLetter = None

root = tk.Tk()

root.geometry("850x150")
root.title("Dynamic Letter Generator")

def doJob():
    printer_name = win32print.GetDefaultPrinter ()
    if sys.version.info >= (3,):
        raw_data = bytes(finishedLetter, "utf-8")
    else:
        raw_data = finishedLetter
    hPrinter = win32print.OpenPrinter(printer_name)
    try:
      hJob = win32print.StartDocPrinter(hPrinter, 1, ("test of raw data", None, "RAW"))
      try:
        win32print.StartPagePrinter(hPrinter)
        win32print.WritePrinter(hPrinter, raw_data)
        win32print.EndPagePrinter(hPrinter)
      finally:
        win32print.EndDocPrinter(hPrinter)
    finally:
      win32print.ClosePrinter(hPrinter)

def firstWindow():

    intro_text = "In this program you will insert a template letter in which fully CAPITALIZED words will be replaced with something from a" \
                 "comma-seperated list or an excel sheet."
    option = "Which would you like to use as an input?"

    topframe = tk.Frame(root)
    topframe.pack(side = tk.TOP, fill = tk.X, anchor = tk.N)
    botframe = tk.Frame(root)
    botframe.pack(side = tk.BOTTOM, fill = tk.X, expand = 1, anchor = tk.S)
    introtxt = tk.Label(topframe, text=intro_text).pack()
    choice = tk.Label(botframe, text=option).pack()
    excelButton = tk.Button(botframe, text = "Excel", command = lambda: secondWindowExcel(topframe, botframe), width = 50).pack(pady = 10)
    excelButton = tk.Button(botframe, text="Comma-seperated List", command= lambda: secondWindowList(topframe, botframe), width = 50).pack(pady = 10)
    pass

def get_excelsheet(tF, bF):

    imp_file_path = filedialog.askopenfilename()
    df = op.load_workbook(imp_file_path)
    thirdWindow(tF, bF)


def secondWindowExcel(tF,bF):
    root.geometry("850x275")
    for i in tF.winfo_children():
        i.destroy()
    for i in bF.winfo_children():
        i.destroy()
    picpath = "pic.jpg"
    img = ImageTk.PhotoImage(Image.open(picpath))
    directions = "In order to use an excel sheet to generate dynamic letters, there are a few steps to follow: "
    stepOne = "Step 1: The first row of your excel sheet has to contain the capitalized words that will be replaced later."
    stepTwo = "Step 2: Each column should contain the respective values for the capitalized words."
    note = "NOTE: Column lengths should match, if they don't, dynamic letters will only be made until the first mismatch. (Put 'nothing' if you want it to be blank)"
    # PUT IN EXAMPLE PIC BETWEEN NOTE AND STEP 3 #
    stepThree = "Click the button below in order to upload your excel file."

    direction_lbl = tk.Label(tF, text = directions).pack()
    steponelbl = tk.Label(tF, text = stepOne).pack()
    steptwolbl = tk.Label(tF, text=stepTwo).pack()
    notelbl = tk.Label(bF, text=note).pack()
    picturelbl = tk.Label(tF, image = img)
    picturelbl.image = img
    picturelbl.pack(side = tk.BOTTOM, expand = 1)
    stepthreelbl = tk.Label(bF, text=stepThree).pack()
    uploadButton = tk.Button(bF, text = "Upload Excel File",width = 50, command = lambda: get_excelsheet(tF,bF)).pack()
    pass

def secondWindowList(tF,bF):
    for i in tF.winfo_children():
        i.destroy()
    for i in bF.winfo_children():
        i.destroy()
    pass

def thirdWindow(tF, bF):
    for i in tF.winfo_children():
        i.destroy()
    for i in bF.winfo_children():
        i.destroy()

    pass

def fourthWindow():
    pass

firstWindow()
root.mainloop()