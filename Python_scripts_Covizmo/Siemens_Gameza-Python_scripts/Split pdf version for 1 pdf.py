import os
from pikepdf import Pdf
import cv2
from PIL import Image
import os
import pandas as pd
import numpy as np
from pathlib import Path
import easygui
import tkinter
import tkinter as tk
from openpyxl import Workbook
from xlrd import open_workbook
import csv
from tkinter import messagebox
import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font

from tkinter.filedialog import askopenfilename
# a dictionary mapping PDF file to original PDF's page range
messagebox.showinfo("showinfo", "Choose file");
filename = tkinter.filedialog.askopenfilename()
path_to_save = filename[:-4]
# name = path_to_save.split("/").str[-1]
# os.mkdir(path_to_save)
messagebox.showinfo("showinfo", "Choose folder");
path_to_folder = tkinter.filedialog.askdirectory()
file2pages = {
    0: [0, 10], # 1st splitted PDF file will contain the pages from 0 to 9 (9 is not included)
    1: [10, 17], # 2nd splitted PDF file will contain the pages from 9 (9 is included) to 11
    2: [17, 24],
    3: [24, 32],  # 1st splitted PDF file will contain the pages from 0 to 9 (9 is not included)
    4: [32, 40],  # 2nd splitted PDF file will contain the pages from 9 (9 is included) to 11
    5: [40, 47],
    6: [47, 53], # 1st splitted PDF file will contain the pages from 0 to 9 (9 is not included)
    7: [53, 60], # 2nd splitted PDF file will contain the pages from 9 (9 is included) to 11
    8: [60, 68],
    9: [68, 77], # 1st splitted PDF file will contain the pages from 0 to 9 (9 is not included)
    10: [77, 84], # 2nd splitted PDF file will contain the pages from 9 (9 is included) to 11
    11: [84, 91],
    12: [91, 99], # 1st splitted PDF file will contain the pages from 0 to 9 (9 is not included)
    13: [99, 106], # 2nd splitted PDF file will contain the pages from 9 (9 is included) to 11
    14: [106, 111]
    # 3rd splitted PDF file will contain the pages from 11 until the end or until the 100th page (if exists)
}

# the target PDF document to split
# filename = "bert-paper.pdf"
# load the PDF file
pdf = Pdf.open(filename)
# make the new splitted PDF files
new_pdf_files = [ Pdf.new() for i in file2pages ]
# the current pdf file index
new_pdf_index = 0
# iterate over all PDF pages
for n, page in enumerate(pdf.pages):
    if n in list(range(*file2pages[new_pdf_index])):
        # add the `n` page to the `new_pdf_index` file
        new_pdf_files[new_pdf_index].pages.append(page)
        print(f"[*] Assigning Page {n} to the file {new_pdf_index}")
    else:
        # make a unique filename based on original file name plus the index
        name, ext = os.path.splitext(filename)
        output_filename = f"{name}-{new_pdf_index}.pdf"
        # save the PDF file
        new_pdf_files[new_pdf_index].save(output_filename)
        print(f"[+] File: {output_filename} saved.")
        # go to the next file
        new_pdf_index += 1
        # add the `n` page to the `new_pdf_index` file
        new_pdf_files[new_pdf_index].pages.append(page)
        print(f"[*] Assigning Page {n} to the file {new_pdf_index}")

# save the last PDF file
name, ext = os.path.splitext(filename)
output_filename = f"{name}-{new_pdf_index}.pdf"
new_pdf_files[new_pdf_index].save(output_filename)
print(f"[+] File: {output_filename} saved.")