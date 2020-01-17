# Author:   Gabriel De Jesus
# Date:     August 17, 2019
# Purpose:  Filter data from Excel spreadsheets
# Simple GUI included for ease of use, but can be run from the command line if needed
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename
import pandas as pd
from pandas import DataFrame
import tkinter as tk

root = Tk(  )

# Process Grant Section
def GrantFilter():
    name = None
    name = askopenfilename(filetypes =(("Excel 97/2003", "*.xls"),("Excel","*.xlsx"),("All Files","*.*")))
    if name != None:
        df_initial = pd.read_excel(name)
        df_final = pd.DataFrame()
        total_rows = len(df_initial.index)
        root.destroy()
        df_initial.columns = df_initial.columns.str.strip()
        # We iterate through rows to match values in a specified column
        # Here, we're looking to match either CA/AZ in State data
        # We can also further specify regions in the state, using zip code
        for row in range(total_rows):
            if ( str(df_initial.loc[row]['State']) == 'CA'):
                if ( int(str(df_initial.loc[row]['Zip'])[:5]) <= 93108 ):
                    df_final = df_final.append(df_initial.loc[row])
        for row in range(total_rows):
            if ( str(df_initial.loc[row]['State']) == 'AZ'):
                df_final = df_final.append(df_initial.loc[row])
        # export data to excel after sorting by any specific column if wanted
        df_final = df_final.sort_values(by=['Dollars Remaining'], ascending=False)
        df_final.to_excel('Grants-Filtered.xlsx', index=False)
    elif name == '':
        root.destroy()
        return
    else:
        root.destroy()
        return

# This function is essentially the same as above, except we can further
# limit any filtering of excel data by implementing an "array" of values
# we want to match, converting it to a set, and then filtering
# the set to an appendable dataframe which we concatenate
# before exporting to a new Excel file.
def EvoFilter():
    evo_name = None
    evo_name = askopenfilename(filetypes =(("Excel 97/2003", "*.xls"),("Excel","*.xlsx"),("All Files","*.*")))
    if evo_name != None:
        #evo_core_df = pd.DataFrame()
        df_evo_initial = pd.read_excel(evo_name, sheet_name="Product Line", skiprows=16)
        evo_core_df = pd.DataFrame()
        det_core_df = pd.DataFrame()
        combined_df = pd.DataFrame()
        total_evo_rows = len(df_evo_initial.index)
        root.destroy()
        evo_core_part_list = ['part1','part2','etc']
        evo_core_part_list_set = set(evo_core_part_list)
        det_core_part_list = ['secondary_part1','etc']
        det_core_part_list_set = set(det_core_part_list)
        #print(df_evo_initial)
        for evo_row in range(total_evo_rows):
            if ( str(df_evo_initial.loc[evo_row]['Product Line'])[:3] in evo_core_part_list_set ):
                evo_core_df = evo_core_df.append(df_evo_initial.loc[evo_row])
        evo_core_df = evo_core_df.sort_values(by=[' YTD 2019 Net Sales'], ascending=False)
        for det_row in range(total_evo_rows):
            if ( str(df_evo_initial.loc[det_row]['Product Line'])[:3] in det_core_part_list_set ):
                det_core_df = det_core_df.append(df_evo_initial.loc[det_row])
        # We can sort data extracted by various attributes that may be in the spreadsheet
        # Total sales, Revenue, Profit, etc.
        det_core_df = det_core_df.sort_values(by=[' YTD 2019 Net Sales'], ascending=False)
        combined_df = combined_df.append(evo_core_df)
        combined_df = combined_df.append(pd.Series(), ignore_index=True)
        combined_df = combined_df.append(det_core_df)

        combined_df.to_excel('Evo-Core-Filtered.xlsx', index=False)
    elif evo_name == '':
        root.destroy()
        return
    else:
        root.destroy()
        return
######
Title = root.title( "Excel Filtering - Gabriel De Jesus")
label = ttk.Label(root, text ="Program by: Gabriel De Jesus. \nNot for distribution.",foreground="red",font=("Times New Roman", 12))
label.pack()

# Menu Bar
#menu = Menu(root)
#root.config(menu=menu)
#file = Menu(menu, tearoff=0)
#file.add_command(label = 'Grant Filtering', command = GrantFilter)
#file.add_command(label = 'EVO Core Filter', command = EvoFilter)
#file.add_command(label = 'Exit', command = lambda:exit())
#menu.add_cascade(label = 'File', menu = file)
# Convenience Buttons
grantbutton = tk.Button(root, text="Grant Filtering (CA/AZ)", command=GrantFilter)
grantbutton.pack()
evobutton = tk.Button(root, text="EVO Filtering", command=EvoFilter)
evobutton.pack()
quitbutton = tk.Button(root, text="Exit", fg="red", bg="black", command=lambda:exit())
quitbutton.pack()
root.mainloop()
