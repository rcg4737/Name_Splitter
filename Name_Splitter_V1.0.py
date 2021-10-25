import tkinter
from tkinter import messagebox, ttk, filedialog  
from ttkthemes import ThemedTk
import pandas as pd
import string
import os



##########   GUI DESIGN   ##########

root = ThemedTk(theme="breeze")
root.title('Name Splitter V1.0')
root.iconbitmap(r"Insert path") #Insert path to company logo

##########   VARIABLES   ##########
def clearApp():
    filePathEntry.delete(0,'end')
    column_drop.delete(0, 'end')
    submitButton["state"] = "enable"

def browse_cmd():
    """Opens file explorer browse dialogue box for user to search for files in GUI."""
    clearApp()
    root.filename = filedialog.askopenfilename()
    filePathEntry.insert(0, root.filename)
    return None


def automation():
    name_file_path = filePathEntry.get()
    column_letter = column_drop.get()
    submitButton["state"] = "disabled"
    
        ### ERROR HANDLING: INVALID PATH OR EMPTY PATH   ####
    if name_file_path == '' or column_letter=='':
        tkinter.messagebox.showerror('Empty File Path','Please select the file that contains the name column and the column letter.')
        clearApp()
        return

    #    READ MAIN FILE
    if os.path.splitext(name_file_path)[1]=='.xlsx':
        
        try:
            name_column = pd.read_excel(name_file_path, usecols=column_letter, engine='openpyxl')
        except:
            tkinter.messagebox.showerror('File Reading Error','Make sure the file is closed and try again.')
            clearApp()
            return
    elif os.path.splitext(name_file_path)[1]=='.csv':
        try:
            name_column = pd.read_csv(name_file_path, usecols=column_letter)
        except:
            tkinter.messagebox.showerror('File Reading Error','Make sure the file is closed and try again.')
            clearApp()
            return
    else:
        tkinter.messagebox.showerror('File Reading Error','The file type is not supported.\n Please contact the developer.')
        clearApp()
        return
    
    name_column= name_column.rename(columns={name_column.columns[0]:"Full_Name"})
    
    split_names = name_column["Full_Name"].str.split(" ", n=0, expand = True) 
    
    name_column = pd.concat([name_column,split_names], axis=1)

    for index, (b, c, d, e, f) in enumerate(zip(name_column[0], name_column[1], name_column[2], name_column[3], name_column[4])):
        if 'ESTATE' in b and 'OF' in c:
            name_column[0][index] = ' '.join([name_column[0][index],name_column[1][index],name_column[2][index]])
            if name_column[4][index]!=None:
                name_column[1][index]=name_column[3][index]
                name_column[2][index]=name_column[4][index]
                name_column[3][index]=''
                name_column[4][index]=''
            else:
                name_column[2][index]=name_column[3][index]
                name_column[1][index]=''
                name_column[3][index]=''
                name_column[4][index]=''
        if b=='THE':
            name_column[0][index] = ' '.join([name_column[0][index],name_column[1][index]])
            if name_column[4][index]!=None:
                name_column[1][index]=name_column[2][index]
                name_column[2][index]=' '.join([name_column[3][index],name_column[4][index]])
                name_column[3][index]=''
                name_column[4][index]=''
            else:
                name_column[1][index]=name_column[2][index]
                name_column[2][index]=name_column[3][index]
                name_column[3][index]=''
        if d==None:
            name_column[2][index]=name_column[1][index]
            name_column[1][index]=''
            name_column[3][index]=''
            name_column[4][index]=''
        if e!=None and f==None:
            name_column[2][index]=' '.join([name_column[2][index],name_column[3][index]])
            name_column[3][index]=''
            name_column[4][index]=''

    for x in name_column.columns:
        name_column[x] = name_column[x].str.replace('.','')
    
    os.chdir(os.path.dirname(name_file_path))
    
    name_column.to_csv('split_names.csv', index=False)
    
    clearApp()


filepathLabel = ttk.Label(root, text="Click the browse button and select the Pay History file.")
filepathLabel.grid(row=0, column=0, pady=10, padx=10)

filePathEntry = ttk.Entry(root, width=50 )
filePathEntry.grid(row=1, column=0, pady=10, padx=10)

browseButton = ttk.Button(root, text='Browse', command= browse_cmd)
browseButton.grid(row=1, column=1, pady=10, padx=10)

column_drop_label = ttk.Label(root, text="Select the letter of the name column.")
column_drop_label.grid(row=4, column=0, pady=10, padx=10)

column_drop = ttk.Combobox(root, values=list(string.ascii_uppercase), width=20)
column_drop.grid(row=5, column=0, padx=10, pady=10)

submitButton = ttk.Button(root, text='Submit', command=automation)
submitButton.grid(row=6, column=0, padx=10, pady=10)

root.mainloop()
