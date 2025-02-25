import pymupdf
import sys
import CleanData
import tkinter
from tkinter import filedialog
from tkinter import messagebox
from tkinter import simpledialog

root = tkinter.Tk()
root.withdraw()

#Change directory to packing slip directory
print("Select your pdf file.")
fPath = filedialog.askopenfilename(title="Select your pdf file")

#If file path exists load it.
if(fPath):
    print("PDF Loaded\n" + fPath)
    doc = pymupdf.open(fPath)
else:
    messagebox.showinfo("Closing","No File Selected")
    sys.exit()

rawPackingList = {}



def parsePDF():
    pageLines = []
    page = 0
    
    for pages in doc: #write to file only relevant raw data
        pgText = pages.get_text("blocks")
        pageLines.append(pgText[56][4].strip()) # "Ordered by" Name
        for line in range(18,20): # Customer Order # and PO #
            pageLines.append(pgText[line][4].strip())
        for line2 in range(66,71): # "Ship To" Details
            if("Address Type" not in pgText[line2][4]):
                pageLines.append(pgText[line2][4].strip())
        for line3 in range(72,len(pgText)-2): # Sku, Internet #, Description, Qty
            if("Internet Number" not in pgText[line3][4]) and "Address Type" not in pgText[line3][4]:
                pageLines.append(pgText[line3][4].strip())
        rawPackingList[page] = pageLines
        page += 1
        pageLines = []
    #process all Ship to store elements
    for list in rawPackingList.values():
        if("Ship to Store" in list[4]):
            list[3] += "-" + list[4].split(" ")[5]
            list.remove(list[4])
        if(len(list) == 12): #combine separated skus and remove dupe
            list[7] += list[8]
            list.remove(list[8])
        elif(len(list) == 14): #remove extra element from end.
            list.remove(list[len(list)-1]) 
        #print(list)


parsePDF()
CleanData.cleanPackingList(rawPackingList)

