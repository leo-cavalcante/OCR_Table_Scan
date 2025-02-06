# python3 -m venv path/to/venv
# source path/to/venv/bin/activate
# python3 -m pip install tk
# python3 -m pip install tkinter
# python3 -m pip install matplotlib 
# python3 -m pip install numpy
# python3 -m pip install pandas 
# python3 -m pip install os 
# python3 -m pip install pillow
# python3 -m pip install pytesseract 
# python3 -m pip install tesseract-ocr
# python3 -m pip install datetime

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

def optimizeDf(old_df: pd.DataFrame) -> pd.DataFrame:
    df = old_df[["left", "top", "width", "text"]]
    df['left+width'] = df['left'] + df['width']
    df = df.sort_values(by=['top'], ascending=True)
    df = df.groupby(['top', 'left+width'], sort=False)['text'].sum().unstack('left+width')
    df = df.reindex(sorted(df.columns), axis=1).dropna(how='all').dropna(axis='columns', how='all')
    df = df.fillna('')
    return df

def mergeDfColumns(old_df: pd.DataFrame, threshold: int = 30, rotations: int = 10) -> pd.DataFrame:
  df = old_df.copy()
  for j in range(0, rotations):
    new_columns = {}
    old_columns = df.columns
    i = 0
    while i < len(old_columns):
      if i < len(old_columns) - 1:
        # If the difference between consecutive column names is less than the threshold
        if any(old_columns[i+1] == old_columns[i] + x for x in range(1, threshold)):
          new_col = df[old_columns[i]].astype(str) + df[old_columns[i+1]].astype(str)
          new_columns[old_columns[i+1]] = new_col
          i = i + 1
        else: # If the difference between consecutive column names is greater than or equal to the threshold
          new_columns[old_columns[i]] = df[old_columns[i]]
      else: # If the current column is the last column
        new_columns[old_columns[i]] = df[old_columns[i]]
      i += 1
    df = pd.DataFrame.from_dict(new_columns).replace('0', '').replace(0, '').replace(np.nan, '')
  return df.dropna(axis='columns', how='all').drop_duplicates().reset_index(drop=True)
  # df = pd.DataFrame.from_dict(new_columns).replace('', np.nan).dropna(axis='columns', how='all')
  # return df.replace(np.nan, '')

def mergeDfRows(old_df: pd.DataFrame, threshold: int = 15) -> pd.DataFrame:
    new_df = old_df.drop_duplicates().iloc[:1]
    for i in range(1, len(old_df)):
        # If the difference between consecutive index values is less than the threshold
        #if abs(old_df.index[i] - old_df.index[i - 1]) < threshold: 
        #    new_df.iloc[-1] = new_df.iloc[-1].astype(str) + old_df.iloc[i].astype(str)
        #else: # If the difference is greater than the threshold, append the current row
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

def upload_file():
    f_types = [('Jpg Files', '*.jpg', '*.png', '*.webp'),
    ('PNG Files','*.png')]   # type of files to select 
    filename = tk.filedialog.askopenfilename(multiple=True,filetypes=f_types)
    col=1 # start from column 1
    row=3 # start from row 3 
    for f in filename:
        img=Image.open(f) # read the image file
        img=img.resize((100,100)) # new width & height
        img=ImageTk.PhotoImage(img)
        e1 =tk.Label(my_w)
        e1.grid(row=row,column=col)
        e1.image = img # keep a reference! by attaching it to a widget attribute
        e1['image']=img # Show Image  
        if(col==3): # start new line after third column
            row=row+1# start wtih next row
            col=1    # start with first column
        else:       # within the same row 
            col=col+1 # increase to next column

## Configuring the Tesseract
os.environ['TESSDATA_PREFIX'] = '/Users/leonardo.cavalcante/Documents/Education - Learning/GitHub/OCR_Table_Scan/Languages/'

    # To check what languages that pytesseract detected used the following print statement:
    # print(pytesseract.get_languages())

special_config = '--psm 12 --oem 1'
languages_ = "eng+fra" # For multiple language use like "eng+fra+spa+deu+chi_sim" and so on

## Uploading photo
my_w = tk.Tk()
my_w.geometry("410x300")  # Size of the window 
my_w.title('www.plus2net.com')
my_font1=('times', 18, 'bold')
l1 = tk.Label(my_w,text='Upload Files & display',width=30,font=my_font1)  
l1.grid(row=1,column=1,columnspan=4)
b1 = tk.Button(my_w, text='Upload Files', 
   width=20,command = lambda:upload_file())
b1.grid(row=2,column=1,columnspan=4)

my_w.mainloop()  # Keep the window open

    # Any image format will do
image_path = "//Images_to_Treat/Image.webp"

    # You can use opencv for that option too
img_pl=PIL.Image.open(image_path)

data = pytesseract.image_to_data(
                        img_pl, 
                        lang=languages_, 
                        output_type='data.frame', 
                        config=special_config)


## Treating the data
data_imp_sort = optimizeDf(data)
df_new_col = mergeDfColumns(data_imp_sort)
merged_row_df = mergeDfRows(df_new_col)

## Cleaning & Outputing the data
cleaned_df = clean_df(merged_row_df.copy())
cleaned_df.reset_index(drop=True, inplace=True)
datetime_str = datetime.today().strftime('%Y-%m-%d at %H.%M.%S')
cleaned_df.to_excel(datetime_str + '.xlsx', index=False, sheet_name='Scanned')