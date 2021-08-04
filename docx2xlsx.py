from tkinter import *
from tkinter.filedialog import askopenfilename
from docx import Document
from os import listdir
import pandas as pd
import os
import zipfile
import csv

global ls
ls=[]

# functions
## get the directory
def getfiledir():
    Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file
    return filename

## click function
def click():
    entered_text=getfiledir()
    ls.append(entered_text)
    outbox.insert(END, entered_text + '\n')

## magic
def exportxls():
    global df

    counter = 0

    for i in ls:
        document = Document(i)
        table = document.tables[0]
        data = [[cell.text for cell in row.cells] for row in table.rows]
        df_temp = pd.DataFrame(data)
        df_temp2 = pd.DataFrame(data) #unflipped

        #df_temp = df_temp.rename(columns=df_temp.iloc[0]).drop(df_temp.index[0]).reset_index(drop=True)
        #df_temp2 = df_temp2.rename(columns=df_temp2.iloc[0]).drop(df_temp2.index[0]).reset_index(drop=True)

        # flip df
        df_temp.index = [0] * len(df_temp)
        df_temp = df_temp.pivot(index=None, columns=0, values=1)

        if counter == 0:
            df = df_temp
        else: #next iteration, append answers to df
            ls_ans_final = []
            ls_ans = df_temp2[1]
            for x in ls_ans:
                ls_ans_final.append(x)

            df_length = len(df)
            df.loc[df_length] = ls_ans_final

        counter+=1

    print(df)


    window.destroy()
    exit()

# main
window = Tk()
window.title("Select files")

## output box - the directory of the selected files will be displayed here
outbox = Text(window, width=50, height=6, wrap=WORD, background="white")
outbox.grid(row=1, column=0, columnspan=2, sticky=W)

## add file button
Button(window, text="add file", width=8, command=click) .grid(row=2, column=0, sticky=W)

## export to excel
Button(window, text="export", width=8, command=exportxls) .grid(row=2, column=1, sticky=W)

# run everything
window.mainloop()
