from tkinter import *
from tkinter.filedialog import askopenfilename
from docx import Document
from os import listdir
import pandas as pd
import os
import zipfile
import csv
import re

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

    #this is also a number of rows excluding the questions row
    counter = 0

    ls_Qs = []
    ls_As = []

    chars = ""
    
    for i in ls:

        document = Document(i)
        table = document.tables[0]
        data = [[cell.text for cell in row.cells] for row in table.rows]
        df_temp = pd.DataFrame(data)
        
        #temp list for questions and answers
        ls_q_temp = df_temp[0] 
        ls_a_temp = df_temp[1]
        
        #add all elements to its own list
        #insert delimeter if needed
        if counter == 0 :
            for j in ls_q_temp:
                #replace commas with '/' and '\n' with spaces so that .csv file doesnt go crazy
                j = j.replace('\n'," ")
                j = j.replace(','," / ")
                ls_Qs.append(j)
                    
        for k in ls_a_temp:
            #replace commas with '/' and '\n' with spaces so that .csv file doesnt go crazy
            k = k.replace('\n'," ")
            k= k.replace(','," / ")
            ls_As.append(k)
            
        counter+=1    
    
    #loop from 1 to the number of files provided
    #export to csv file with right formatting hopefully
    pos_count=0
    ls_2write=[]
    
    #if you are new user, change this 'C:\\Users\Dave.comia\Documents\Python Scripts\' to the directory you want to save file in
    with open('C:\\Users\Dave.comia\Documents\Python Scripts\sample.csv','w', newline='', encoding='utf-8-sig') as csvfile:
        writer=csv.writer(csvfile, quotechar=' ')    
        
        for e in range(1, counter+1):
            #export headers
            if e == 1:
                writer.writerow(ls_Qs)

            #export contents of ls_As from pos(1) to pos(l en(ls_Qs)) at index(e+1)
            #pos = position in list
            #index = row number
            ls_2write = ls_As[pos_count:pos_count+len(ls_Qs)]
            pos_count = pos_count + len(ls_Qs)
            writer.writerow(ls_2write)
    
    print(ls_Qs)
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
