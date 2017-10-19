import csv
import xlrd
from os import listdir, chdir, _exit
from os.path import isfile, join
import time

files = [f for f in listdir("toExport/") if isfile(join("toExport/", f))]
print("Exporting", len(files), "files.")

start = time.time()
print("Convertion started at:", round(start), "epochs")

average = 0.0
count = 0
while True:
    count += 1
    start = time.time()
    filecount = 0
    for f in files:
        filecount += 1
        if ".xlsm" in f:
            doc = xlrd.open_workbook("toExport/" + f)
            print(doc.sheet_names())
            sheet = doc.sheets()[1]
            opened = open("exported/" + f.replace(".xlsm", "-export-") + sheet.name() + ".csv", "w")
            print(filecount, "Exporting:", f)
            writer = csv.writer(opened)
            for row in range(0, sheet.nrows):
                out = []
                for cell in sheet.row_values(row):
                    out.append(str(cell).encode('ascii', 'ignore').decode())
                writer.writerow(out)
            sheet = doc.sheets()[3]
            for row in range(0, sheet.nrows):
                out = []
                for cell in sheet.row_values(row):
                    out.append(str(cell).encode('ascii', 'ignore').decode())
            opened.close()
            print("Time elapsed:", round(time.time() - start, 2), "/s")
    average += round(time.time() - start, 2)
    if count%5==0:
        average = average / count
    print("Average time taken:", average, "\n\n")
