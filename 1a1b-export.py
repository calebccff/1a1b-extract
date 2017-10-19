"""
Need to fix output not saving
manual entry of cells
"""

from csv import reader
from os import listdir, chdir
from os.path import isfile, join
import time, xlrd, csv

from tkinter import *

import os

# Tk init
from tkinter import messagebox, filedialog

WIDTH, HEIGHT = 400, 600

root = Tk()
root.resizable(width=False, height=False)
root.geometry('{}x{}'.format(WIDTH, HEIGHT))

folder_directory = ""
cells_file = ""
cells = []
output = list(reader(open("template.csv", "r")))

#Event handlers to clear text box
def on_entry_click(event):
    """function that gets called whenever entry is clicked"""
    if cells_entry.get() == 'col:row':
        cells_entry.delete(0, "end") # delete all the text in the entry
        cells_entry.insert(0, '') #Insert blank for user input
        cells_entry.config(fg='black')
def on_focusout(event):
    if cells_entry.get() == '':
        cells_entry.insert(0, 'col:row')
        cells_entry.config(fg='grey')

def file_open():
    global folder_directory
    folder_directory = filedialog.askdirectory()
    print(folder_directory)
    folder_label.grid_remove()
    button_open.grid_remove()
    cells_label.grid(row=0, column=0)
    cells_load.grid(row=1, column=0)
    cells_enter.grid(row=1, column=1)


def cells_open():
    global cells_file, cells
    cells_enter.grid_remove()
    cells_file = filedialog.askopenfilename()
    print(cells_file)
    cells = open(cells_file, "r").read().split("\n")
    for i in range(len(cells)):
        cells[i] = cells[i].split(":")
    run_button.grid(row=5, column=0)

def cell_add():
    global cells
    cells.append(cells_entry.get().split(":"))
    cells_entry.delete(0, "end")
    print(cells)
    if len(cells) < 2:
        run_button.grid(row=7, column=1)

def cells_type():
    global cells_entry
    cells_load.grid_remove()

    cells_entry = Entry(root)
    cells_entry.insert(0, "col:row")
    cells_entry.bind("<FocusIn>", on_entry_click)
    cells_entry.bind('<FocusOut>', on_focusout)
    cells_entry.config(fg="grey")
    cells_entry.grid(row=6, column=0)

    cell_button_add = Button(root, text="Add cell", command=cell_add)
    cell_button_add.grid(row=6, column=1)


def run():
    global output, cells

    start = time.time()
    print("Exraction started at:", round(start), "epochs")

    z = ord("a")

    chdir(folder_directory)
    files = [f for f in listdir() if isfile(join(f))]

    if len(files) == 0:
        print("Please export your 1a reports to the folder '1a1b-data' as a '.csv' file")

    for i in range(len(files) - 1, 0, -1):
        if "xlsm" in files[i]:
            del files[i]

    opened = []
    for i in range(len(files)):
        opened.append(list(reader(open(files[i], "r"))))

    for i in range(len(opened)):
        print("Currently data mining:", files[i])
        print("Time elapsed:", round(time.time() - start, 2), "/s")
        data = []
        xpos, ypos = 0, 0
        for cell in cells:
            if len(cell[0]) == 1:
                xpos = ord(cell[0]) - z
            else:
                parts = [ord(l) - z for l in list(cell[0])]
                xpos = 26 * (parts[0] + 1)
                xpos += parts[1]

            ypos = int(cell[1]) - 1
            data.append(opened[i][ypos][xpos].replace(",", "/"))
        output.append(data)

    save_button.grid(row=13, column=0)


def save():
    save_location = filedialog.asksaveasfilename()
    print(save_location)
    f = open(save_location, "w")
    f.write(str(output).replace("], [", "\n").replace("[", "").replace("]", "").replace("'", ""))
    f.close()
    Label(root, text="DONE!\nThe output file can be located at:\n"+save_location).grid(row=12, column=0)

top = Frame(root)
bottom = Frame(root)

if False:
    messagebox.showinfo("Instructions",
                        "Please enter the column/row location of each cell you wish to data min in the form\n'x y' "
                        "where x is the column letter and y is the row number.\nAlso ensure that they are seperated "
                        "by a single space.\n\nAt the moment no input validation is performed\nso the program WILL "
                        "crash if you don't format your input. When you have finished type 'done' to finish.\n\nIf "
                        "you wish to load the cell adresses from a text file please type 'load' followed by the name "
                        "of the file.")


folder_label = Label(root, text="Open the folder containing your exported 1a1b csv files")
button_open = Button(root, text="Open 1a1b directory", command=file_open)

cells_label = Label(root, text="Load cell addresses from a file or enter manually?")
cells_load = Button(root, text="Load", command=cells_open)
cells_enter = Button(root, text="Enter manually", command=cells_type)

run_button = Button(root, text="RUN!", command=run)
save_button = Button(root, text="Save as", command=save)

folder_label.grid(row=0, column=0)
button_open.grid(row=1, column=0)

root.mainloop()

"""ORDER
file_open() - open a1ab directory
cells_open() - open cells.txt
cells_type() - enter cells manually
run()
save()



"""