
#Imports the libraries
import zipfile
from os import listdir, chdir
from os.path import isfile, join
import time, xlrd, csv
import os
import tkinter as tk
from tkinter import filedialog
from fnmatch import fnmatch

root = tk.Tk()
root.withdraw()

#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!
"""
In case of a crash with the exception "CompDocError: Workbook corruption: seen[2] == 4" or similar
Open the file CompDoc.py under your python installation directory (The error should tell you where the file is)
Find and comment out lines 160 and 161 by inserting a '#' at the beginning of both lines
This is an error with loading old Excel documents (pre 2010)
"""
#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!

#!#!
"""
TODO:

 -- Implement header loading (load from several text files containing headers)
 -- Load multiple cells files
 -- Testing - Black box
            - StabilityEED FIX ~ Some cells are "out of range" apparently
 -- Clean up code
 -- Finish directory management. DONE
"""
#!#!

print("1a1b directory can contain subfolders, if you wish to exclude a folder make sure the folder name contains\n"
      " the string 'EXCLUDE'.")
print("\nThe exported database file is the csv file where all the data will be placed, file location must use '/' to\n"
      "seperate folders. You do not need to include a file extension.")
print("\nThe headers file is a text file containing the headers for the database, each one should be on a new line\n"
      "that correlates with the lines in the cell addresses file.")
print("\nThe cell addresses file contains a list of every cell you want to take data from, it should be in the format\n"
      "'col:row' where the column is the letter in Excel, e.g. 'af:15'. ALL CELL ADDRESSES MUST BE LOWERCASE")
print("\nAll files containing cell adresses should be placed in the same directory as this script and start with the \n"
      "word 'cells', all files containing headers for the output file should start with the word 'headers'\n\n\n")

#Declaring some variables
#########################

#This contains the names of all useful sheets in a given workbook
USEFUL_SHEETS = ["Attachment 1a", "1a", "Attachment 1b", "1b", "Attachment 1c", "1c", "Project Progress"]

WORKINGDIR = os.getcwd()

print("Enter the location of the directory containing 1a1bs ->")
ROOT = filedialog.askdirectory()+"/" #Directory of files to export
print("Please browse to the folder you want to use for the output")
EXPORT = filedialog.askdirectory()+"/"
print(EXPORT)
DATABASE = EXPORT+"database.csv"

HEADERS = []
CELLS = []
#Init headers
for fname in os.listdir(WORKINGDIR):
    try:
        tempf = open(fname, "r", newline="")
        if len(fname) > 8 and "template" == fname[:8]:
            HEADERS.append(tempf)
        elif len(fname) > 5 and "cells" == fname[:5]:
            CELLS.append(tempf)
    except IsADirectoryError:
        pass


pattern = ".xls" #File names must contain that string
files = [] #List of filenames/directories
#sheets = ["1a", "Attachment 1a", "1b", "Attachment 1b", "Progress Report", ] #What to export from each excel document

for path, subdirs, fs in os.walk(ROOT):  # Gets all the files from a directory, includes subdirs (must match pattern)
    for name in fs:
        if pattern in name and "csv" not in name and "EXCLUDE" not in name:
            files.append(os.path.join(path, name))

prog_start = time.time() #Get the time since the epoch since program started
print("Extraction started at", round(prog_start), "epochs") #LOGS

def export(): #Contains code to convert from xls(m) to csv
    filecount = 0 #How many files have been converted
    averages = [] #Used to measure average time for one file to export
    
    def export_file(f): #Given a file f this function exports that file as a csv, this takes the longest ~1.3 seconds per file
        ext = ".xlsm"
        if "xlsm" not in f:
            ext = ".xls"
        try:
            xl_doc = xlrd.open_workbook(f) #Catch errors with corrupt/inaccessible files
        except: #Not excepting a specific error because xlrd can throw quite a few.
            print("Skipping", f, "Couldn't read file")
            return False
        fileName = f.split("/")[-1]
        for sheet in xl_doc.sheets(): #Iterate through the sheets in the excel doc
            if sheet.name in USEFUL_SHEETS:
                opened = open(EXPORT+fileName.replace(ext, "-export-") + sheet.name + ".csv", "w") #Open the csv file for writing
                writer = csv.writer(opened) #Create a new CSV writer object
                for row in range(sheet.nrows): #Iterate through the rows of the sheet
                    out = [] #Create a list for storing cells
                    for cell in sheet.row_values(row): #Iterate through the cells in the row
                        out.append(str(cell).encode('ascii', 'ignore').decode()) #Decode each cell and append
                    writer.writerow(out) #Write each row to the file
                opened.close() #Close the file
                print("Exported sheet:", sheet.name)  # LOGS
        return True
        
    for f in files: #Iterate through each file in the index
        print("Exporting", f) #LOGS
        ext = ".xlsm"
        if "xlsm" not in f:
            ext = ".xls"
        filetime = time.time() #Time at which we started exporting this file
        if export_file(f): #Export it!
            filecount += 1 #We've got another one!
            print("Exported", filecount, "of", len(files), "documents") #MORE LOGS!!!1!!
            m, s = divmod(round(time.time() - prog_start, 2), 60) #Some clever mathy stuff to convert seconds into human time
            h, m = divmod(m, 60)

            print("Time elapsed   :", "%02d:%02d:%02d" % (h, m, s)) #Display the human readable time (logs...)
            averages.append(time.time() - filetime) #Calculate the average time to export 1 file
            average = round(sum(averages) / len(averages), 2)
            if len(averages) > 500: #Average is based on the last 500 files!
                del averages[0]
            print("Average time   :", average, "/seconds") #Tell the user more stuffs
            m, s = divmod(average * (len(files) - filecount), 60) #MAths again
            h, m = divmod(m, 60)
            print("Time remaining :", "%02d:%02d:%02d" % (h, m, s), "\n") #Ooooh an ETR!
    
    print("Successfully converted", len(files), "files to CSV in", round(time.time()-prog_start, 2), "seconds") #FINSISHED


def extract(): #The fun bit (surprisingly less intensive, about 10 seconds for EVERYTHING)
    files = [] #We need to index all the CSV files this time
    for path, subdirs, fs in os.walk(EXPORT):  # Gets all the files from a directory, includes subdirs (must match pattern)
        for name in fs:
            if ".csv" in name: # Here I'm assuming that every CSV file in the folder is one that we want, this could cause errors
                files.append(os.path.join(path, name)) #Add each file to the list
    z = ord("a") #Used for accessing cell adresses, ord converts a character to its
                 #ASCII representation so z is used a baseline
                 #ord("c")-z would return the number of columns across a cell is where is 'a' is 1

    database = open(DATABASE, "w", newline="") #Open the DATABASE file, this is where the info ends up
    writer = csv.writer(database) #Make a new CSV writer to write to it

    #for tempf in HEADERS:

    headers = (tempf.readlines() for tempf in HEADERS) #Grab the column headers from a template
    print(headers)
    headersStr = ""

    for h in headers:
        headersStr += h.replace("\n", "")+","
    headers = headersStr
    writer.writerow(headers.split(",")) #Write the headers to to the file

    for i in range(len(files) - 1, 0, -1): #Iterate backwards through the files and delete any non-csv files from the list
        if not ".csv" in files[i] or not "-1a" in files[i]:
            del files[i] # Pythonic syntax!

    filedata = [] #The contents of the files
    for i in range(len(files)):
        print(files[i])
        filedata.append(list(csv.reader(open(files[i], "r")))) #Yup just load ALL the files into RAM

    cells = (tempf.readlines() for tempf in HEADERS) #Get a list of useful cells

    for i in range(len(cells)): #Turn the list into a NICE format
        cells[i] = cells[i].replace("\n", "").split(":")
    for i in range(len(filedata)): #Go through each file
        print("Extracting", files[i]) #EXTRACT that data
        out = [] #Each row get's stored here before being written to the disk

        xpos, ypos = 0, 0 #Cell addresses
        for cell in cells: #Each cell to look at
            if len(cell[0]) == 1: #EZ column headers
                xpos = ord(cell[0]) - z
            else: #Pesky column headers with two letters
                parts = [ord(l) - z for l in list(cell[0])] #ORD() each part of the header, multiply the first by 26
                xpos = 26 * (parts[0] + 1)
                xpos += parts[1]

            ypos = int(cell[1]) - 1 #Get the row (PYTHON is 0 based to take 1 away)
            try:
                out.append(filedata[i][ypos][xpos].replace(",", "/")) #Add each cell to the list and get rid of any ',' so as not to confuse the writer
            except IndexError as e: #No longer crashes.
                print("ERROR:", e)
                print("AT:", cell, "-", xpos, ":", ypos)
        writer.writerow(out) #Write the row
    database.close() #Once we've written everything to the database we have to close it

#First code to actually be executed
export() #Export the files to CSV
print("Attempting to extract...") #LOGGING
extract() #Extract the data from all the CSVs

