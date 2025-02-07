## PREPARING THE FIELD
  # python3 -m venv path/to/venv
  # source path/to/venv/bin/activate
  # python3 -m pip install customtkinter
  # python3 -m pip install datetime
  # python3 -m pip install matplotlib 
  # python3 -m pip install numpy
  # python3 -m pip install openpyxl
  # python3 -m pip install os 
  # python3 -m pip install pandas 
  # python3 -m pip install pillow
  # python3 -m pip install pytesseract 
  # python3 -m pip install random
  # python3 -m pip install shutil
  # python3 -m pip install string
  # python3 -m pip install tesseract-ocr
  # python3 -m pip install tk       # python3 -m pip install tkinter
  # python3 -m main.py

import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter.filedialog import askopenfile
from PIL import Image, ImageTk

import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import PIL
import pytesseract
import os
from datetime import datetime

import customtkinter
from tkinter import *
from tkinter import filedialog
import tkinter as tk
from PIL import ImageTk, Image
import os
import shutil
import random  
import string
from tkinter import messagebox
from pathlib import Path

## PICTURE UPLOAD -- FUNCTIONS

def setPreviewPic(filepath):
    global img
    img = Image.open(filepath)
    img = img.resize((250,250), Image.LANCZOS)
    img = ImageTk.PhotoImage(img)
    lbl_show_pic = tk.Label(frame, bg='#1F6AA5',image=img)
    lbl_show_pic.grid(row=1, column=0, columnspan=3,pady=5,ipady=0,sticky="nswe")
    pathEntry.insert(0, filepath)

def get_num_pixels(filepath):
    # global horizontal, vertical
    width, height = Image.open(filepath).size

    vertical = height
    horizontal = width // 5

    # if width > 1.5 * height:
    #     vertical = height // 10
    #     horizontal = width // 20
    #
    # elif width < 1.5 * height:
    #     vertical = height // 20
    #     horizontal = width // 10
    #
    # else:
    #     vertical = height // 10
    #     horizontal = width // 10

    return horizontal, vertical

def selectPic():
    global filename, horizontal, vertical

    filename = filedialog.askopenfilename(
        initialdir=os.getcwd(), 
        title="Select Image",
        # filetypes=[("images files","*.png *.jpg *.jpeg *.webp"),] --> OLD CODE
        filetypes=(("png images","*.png"), ("jpg images", "*.jpg"), ("jpeg images", "*.jpeg"), ("webp images","*.webp"))
    )
    setPreviewPic(filename)

    horizontal, vertical = get_num_pixels(filename)

def OutputFile():
    # Any image format will do
    image_path = filename






    # You can use opencv for that option too
    img_pl=PIL.Image.open(image_path)

    data = pytesseract.image_to_data(
                        img_pl, 
                        lang=languages_, 
                        output_type='data.frame', 
                        config=special_config)

    ## Treating the table data
    data_imp_sort = optimizeDf(data)
    df_new_col = mergeDfColumns(data_imp_sort)
    # df_new_col = mergeDfColumns(data_imp_sort, horizontal, horizontal)
    merged_row_df = mergeDfRows(df_new_col)
    # merged_row_df = mergeDfRows(df_new_col, vertical)

    ## Cleaning & Outputing the data
    cleaned_df = clean_df(merged_row_df.copy())
    cleaned_df.reset_index(drop=True, inplace=True)

    ## CHANGING THE WORKING DIRECTORY TO SAVE FILE IN DOWNLOADS
    desktop_path = str(Path.home() / "Desktop")
    os.chdir(desktop_path)
    # print(os.getcwd())

    ## SAVE COPY OF PHOTO IN DOWNLOADS
    # shutil.copy(filename, f"{datetime_str}.{fileformat_from_filename_splitted[-1]}")
    setPreviewPic('Image.webp')

    ## PREPARING THE NAME OF THE FILE
    fileformat_from_filename_splitted = filename.split('.')
    # randomText = ''.join((random.choice(string.ascii_lowercase) for x in range(12)))
    # shutil.copy(filename, f"{randomText}.{fileformat_from_filename_splitted[-1]}")
    datetime_str = datetime.today().strftime('%Y-%m-%d at %H.%M.%S')

    ## SAVING THE FILE
    # datetime_str = datetime.today().strftime('%Y-%m-%d at %H.%M.%S')
    cleaned_df.to_excel(datetime_str + '.xlsx', index=False, sheet_name='Scanned')

    messagebox.showinfo("Success", "Table Scanned Successfully")

## TABLE SCAN -- FUNCTIONS

def optimizeDf(old_df: pd.DataFrame) -> pd.DataFrame:
    df = old_df[["left", "top", "width", "text"]]
    df['left+width'] = df['left'] + df['width']
    df = df.sort_values(by=['top'], ascending=True)
    df = df.groupby(['top', 'left+width'], sort=False)['text'].sum().unstack('left+width')
    df = df.reindex(sorted(df.columns), axis=1).dropna(how='all').dropna(axis='columns', how='all')
    df = df.fillna('')
    return df

def mergeDfColumns(old_df: pd.DataFrame, threshold: int = 125,
                   rotations: int = 75) -> pd.DataFrame:
    df = old_df.copy()
# def mergeDfColumns(old_df: pd.DataFrame, threshold: int = 75,
#                    rotations: int = 30) -> pd.DataFrame:
#    df = old_df.copy()
    for j in range(0, rotations):
        new_columns = {}
        old_columns = df.columns
        i = 0
        while i < len(old_columns):
            if i < len(old_columns) - 1:
                # If the difference between consecutive column names is less than the threshold
                if any(old_columns[i + 1] == old_columns[i] + x for x in range(1, threshold)):
                    new_col = df[old_columns[i]].astype(str) + df[old_columns[i + 1]].astype(str)
                    new_columns[old_columns[i + 1]] = new_col
                    i = i + 1
                else:  # If the difference between consecutive column names is greater than or equal to the threshold
                    new_columns[old_columns[i]] = df[old_columns[i]]
            else:  # If the current column is the last column
                new_columns[old_columns[i]] = df[old_columns[i]]
            i += 1
        df = pd.DataFrame.from_dict(new_columns).replace('0', '').replace(0, '').replace(np.nan, '')
    return df.dropna(axis='columns', how='all').drop_duplicates().reset_index(drop=True)
    # df = pd.DataFrame.from_dict(new_columns).replace('', np.nan).dropna(axis='columns', how='all')
    # return df.replace(np.nan, '')

def mergeDfRows(old_df: pd.DataFrame, threshold: int = 10) -> pd.DataFrame:
    new_df = old_df.drop_duplicates().iloc[:1]
    for i in range(1, len(old_df)):
        # If the difference between consecutive index values is less than the threshold
        # if abs(old_df.index[i] - old_df.index[i - 1]) < threshold:
        #    new_df.iloc[-1] = new_df.iloc[-1].astype(str) + old_df.iloc[i].astype(str)
        # else: # If the difference is greater than the threshold, append the current row
        # new_df = new_df.append(old_df.iloc[i])
        # df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        # new_df = pd.concat([new_df, old_df], ignore_index=True).replace(np.nan, '').replace(0, '').drop_duplicates()
        new_df = pd.concat([new_df, old_df], ignore_index=True).replace('0', '').replace(0, '').replace('', np.nan)
    return new_df.drop_duplicates().dropna(axis='rows', how='all').reset_index(drop=True)

def clean_df(df):
    # Remove columns with all cells holding the same value and its length is 0 or 1
    df = df.loc[:, (df != df.iloc[0]).any()]
    # Remove rows with empty cells or cells with only the '|' symbol
    df = df[(df != '|') & (df != '') & (pd.notnull(df))]
    # Remove columns with only empty cells
    df = df.dropna(axis=1, how='all')
    return df.fillna('')


## Configuring the Tesseract
os.environ['TESSDATA_PREFIX'] = '/Users/leonardo.cavalcante/Documents/Education - Learning/GitHub/OCR_Table_Scan/Languages/'
# os.environ['TESSDATA_PREFIX'] = os.getcwd() + '/Languages/'
# print(os.getcwd() + '/Languages/')

  # To check what languages that pytesseract detected used the following print statement:
  # print(pytesseract.get_languages())

special_config = '--psm 12 --oem 1'
languages_ = "eng" # For multiple language use like "eng+fra+spa+deu+chi_sim" and so on



## CHANGING THE WORKING DIRECTORY TO DESKTOP or DOWNLOADS
desktop_path = str(Path.home() / "Desktop")
# downloads_path = str(Path.home() / "Donwloads")
os.chdir(desktop_path)
# print(os.getcwd())

## SETTING THE APP
customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("blue")
app = customtkinter.CTk()
app.title("Upload Image of Table to be Scanned")
app.geometry("480x380")

frame=customtkinter.CTkLabel(app,text="")
frame.grid(row=0,column=0,sticky="w",padx=50,pady=20)

## App buttons
selectBtn=customtkinter.CTkButton(frame,text="Browse Image",width=50,command=selectPic)
# horizontal, vertical = get_num_pixels(filename)
pathEntry=customtkinter.CTkEntry(frame,width=200)
saveBtn=customtkinter.CTkButton(frame, text="Scan Table", width=50, command=OutputFile)
lbl_show_pic = tk.Label(frame, bg='#1F6AA5')

setPreviewPic("Image.webp")

selectBtn.grid(row=0,column=0,padx=1,pady=5,ipady=0,sticky="e")
pathEntry.grid(row=0,column=1,padx=1,pady=5,ipady=0,sticky="e")
saveBtn.grid(row=0,column=2,padx=1,pady=5,ipady=0,sticky="e")
lbl_show_pic.grid(row=1, column=0, columnspan=3,pady=5,ipady=0,sticky="nswe")

app.resizable(False, False)
app.mainloop()