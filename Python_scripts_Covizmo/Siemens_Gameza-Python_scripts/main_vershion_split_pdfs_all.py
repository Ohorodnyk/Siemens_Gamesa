import os
from fnmatch import fnmatch
import subprocess
import tkinter
from datetime import time

import camelot
import tabula
import pandas as pd
import numpy as np
import PyPDF2
import fitz
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import csv

import xlsxwriter
from camelot import read_pdf
import os
from tabula import read_pdf
from tabulate import tabulate
import pandas as pd
import io
import cv2
from PIL import Image
from pikepdf import Pdf

from tkinter.filedialog import askopenfilename
messagebox.showinfo("showinfo", "Choose folder");
foldername = tkinter.filedialog.askdirectory()
messagebox.showinfo("showinfo", "Choose folder to save result files");
foldername_to_save = tkinter.filedialog.askdirectory()
messagebox.showinfo("showinfo", "Choose folder to save fotos");
foldername_to_save_foto = tkinter.filedialog.askdirectory()
pattern = "*.pdf"
directory = "final"
for path, subdirs, files in os.walk(foldername):
    for file in files:
        if fnmatch(file, pattern):
            try:
                if file:
                    # Open the PDF File
                    path_to_save = os.path.join(path, file)
                    pdf_file = PyPDF2.PdfFileReader(path_to_save)
                    print(pdf_file)

                # Name of input file
                # a = os.path.basename(file);

                # path = os.path.join(foldername_to_save + "/", directory)
                # os.mkdir(path)

                # ------------- read tables and save CSV, XSLX
                # Read a PDF File
                df = tabula.read_pdf(path_to_save, pages='all')[0]

                #  --------- read and save IMAGES
                num = 1
                idx = 0
                splt = np.array([])
                doc = fitz.open(path_to_save)
                for i in range(len(doc)):
                    for img in doc.get_page_images(i):
                        xref = img[0]
                        pix = fitz.Pixmap(doc, xref)
                        if (pix.n < 5):  # this is GRAY or RGB
                            # if ( i%2==0 ):
                            pix.save(foldername_to_save_foto + "/p%s-%s.png" % (i, xref))  # -- save all images into folder
                            im = cv2.imread(foldername_to_save_foto + "/p" + str(i) + "-" + str(xref) + ".png")
                            img = Image.open(foldername_to_save_foto + "/p" + str(i) + "-" + str(xref) + ".png")
                            if ((img.height) > 1200):
                                splt = np.append(splt, int(i))
                            height = img.height
                """
                print(splt)
                if len(splt) > 0:
                    for pp in range(len(splt) - 1):
                      pdf_writer = PdfFileWriter()
                      for page in range(int(splt[pp]), int(splt[pp]+1)):
                            pdf_writer.addPage(pdf_file.getPage(page))
                      filename_splt = foldername + "/Split" + str(pp) + ".pdf"
                      with open(filename_splt, 'wb') as file1:
                        pdf_writer.write(file1)
                """
                if len(splt) > 1:
                    pdf = Pdf.open(path_to_save)

                    file2pages = {}
                    print(file2pages)
                    for i in range(len(splt) - 1):
                        file2pages[i] = [int(splt[i]), int(splt[i + 1])]
                    print(file2pages)
                    file2pages[len(splt) - 1] = [int(splt[len(splt) - 1]), int(len(doc))]
                    print(file2pages)
                    # file2pages[len(splt)] = [splt[len(splt)], int(len(doc))]
                    # make the new splitted PDF files
                    new_pdf_files = [Pdf.new() for i in file2pages]
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
                            name, ext = os.path.splitext(path_to_save)
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
                    name, ext = os.path.splitext(path_to_save)
                    output_filename = f"{name}-{new_pdf_index}.pdf"
                    new_pdf_files[new_pdf_index].save(output_filename)
                    print(f"[+] File: {output_filename} saved.")
            except Exception:
                # os.rename(path_to_save2, path_to_save2 + "- YOU NEED TO MAKE IT MANUALLY")
                pass