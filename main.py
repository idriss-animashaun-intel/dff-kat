import pandas as pd
import glob, os
from tkinter import Tk
from tkinter import Button
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from pathlib import Path
import numpy as np
from tkinter import *
from time import sleep
import subprocess
import shutil
import webbrowser

current_directory = os.getcwd()
print(current_directory)

dir = os.path.join("outputs")
if not os.path.exists(dir):
    os.mkdir(dir)

sub_dir = os.path.join("outputs//jmp_outputs")
if not os.path.exists(sub_dir):
    os.mkdir(sub_dir)

modules = ["ARR", "SCN", "PTH", "TPI", "IO", "GT", "FUN"]
sub_modules = ['Master', 'File', 'First']

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    path = Path(__file__).parent / relative_path
    return str(path)

def master_file():
    dumpFolder()
    global df_results
    global dfs
    global atts
    global outfile

    base_path = folder_path+'//'

    df_cor = pd.read_csv(base_path + "CorLot.csv")
    df_ref = pd.read_csv(base_path + "RefLot.csv")

    os.chdir(base_path)
    for file in glob.glob("DFF_Tokens*.csv"):
        dff_tok_path = file

    df_results =pd.read_csv(base_path + dff_tok_path)

    df3 = df_results[['DFF_TokenName', 'ModuleName']]
    df_result_tokens_extracted = df3[df3['DFF_TokenName'].str.contains("_0$", na=False)]
    df_result_tokens_extracted['DFF_TokenName'] = df_result_tokens_extracted['DFF_TokenName'].str.replace("_0$", '')
    

    list_pipe_cols = []
    for i in df_result_tokens_extracted['DFF_TokenName']:
        list_pipe_cols.append(i)

    def pipe_delimination(comp_csv):
        drop_list = []

        for i in range(0,len(comp_csv.columns)):
            if comp_csv.columns[i] in list_pipe_cols:
                temp = []
                collect = df_results['DFF_TokenName'][df_results['DFF_TokenName'].str.contains(comp_csv.columns[i] + '_')]
                for j in collect:
                    temp.append(j)
                comp_csv[temp] = comp_csv[comp_csv.columns[i]].str.split("|",expand=True)
                drop_list.append(comp_csv.columns[i])


        comp_csv = comp_csv.drop(drop_list, 1)
    
    pipe_delimination(df_cor)
    pipe_delimination(df_ref)




    df0 = pd.merge(df_ref, df_cor,  how='outer', left_on=['WAFER_ID','SORT_X','SORT_Y'], right_on = ['WAFER_ID','SORT_X','SORT_Y'])

    df0 = df0.reindex(sorted(df0.columns), axis=1)

    all_cols = df0.columns.values

    col_start = ['LOT_x',
    'LOT_y',
    'OPERATION_x',
    'OPERATION_y',
    'DFF_OPTYPE_x',
    'DFF_OPTYPE_y',
    'FACILITY_x',
    'FACILITY_y',
    'DEVREVSTEP_x',
    'DEVREVSTEP_y',
    'FLOW_STEP_x',
    'FLOW_STEP_y',
    'PROGRAM_NAME_x',
    'PROGRAM_NAME_y',
    'WAFER_ID',
    'SORT_X',
    'SORT_Y',
    'INTERFACE_BIN_x',
    'INTERFACE_BIN_y',
    'FBIN_x',
    'FBIN_y']

    for i in all_cols:
        if i not in col_start:
            col_start.append(i)

    df1 = df0.reindex(columns=col_start)
    rename_col = []

    for i in col_start:
        if '_x' in i:
            rename_col.append(i.replace('_x', '_OLD'))
        elif '_y' in i:
            rename_col.append(i.replace('_y', '_NEW'))
        else:
            rename_col.append(i)

    df1.columns = rename_col
    for col in df1:
        if df1[col].dtypes == 'O':
            df1[col] = df1[col].str.replace("'", '')

    att_df_arr = df3[~df3['ModuleName'].str.contains("DE", na=False) & ~df3['ModuleName'].str.contains("GT$", na=False) & df3['ModuleName'].str.contains("ARR", na=False)]
    att_df_scn = df3[~df3['ModuleName'].str.contains("DE$", na=False) & ~df3['ModuleName'].str.contains("GT$", na=False) & df3['ModuleName'].str.contains("SCN", na=False)]
    att_df_fun = df3[~df3['ModuleName'].str.contains("DE$", na=False) & ~df3['ModuleName'].str.contains("GT$", na=False) & df3['ModuleName'].str.contains("FUN", na=False)]
    att_df_pth = df3[df3['ModuleName'].str.contains("PTH", na=False)]
    att_df_tpi = df3[df3['ModuleName'].str.contains("TPI", na=False) & ~df3['ModuleName'].str.contains("GFX", na=False)]
    att_df_io = df3[df3['ModuleName'].str.contains("CLK", na=False)]
    att_df_gt = df3[df3['ModuleName'].str.contains("GT$", na=False) | df3['ModuleName'].str.contains("DE$", na=False) | df3['ModuleName'].str.contains("GFX", na=False)]

    col_df_arr = []
    col_df_scn = []
    col_df_fun = []
    col_df_pth = []
    col_df_tpi = []
    col_df_io = []
    col_df_gt = []

    cols = [
        col_df_arr,
        col_df_scn,
        col_df_fun,
        col_df_pth,
        col_df_tpi,
        col_df_io,
        col_df_gt
    ]

    atts = [
        att_df_arr,
        att_df_scn,
        att_df_fun,
        att_df_pth,
        att_df_tpi,
        att_df_io,
        att_df_gt
    ]

    for n in range(0,len(cols)):
        for i in df1.columns:
            for j in atts[n]['DFF_TokenName']:
                if j in i:
                    cols[n].append(i)

    start_cols=['LOT_OLD',
    'LOT_NEW',
    'OPERATION_OLD',
    'OPERATION_NEW',
    'WAFER_ID',
    'SORT_X',
    'SORT_Y',
    'FBIN_OLD',
    'FBIN_NEW'
    ]

    dfs = [[],[],[],[],[],[],[]]
    for i in range(0,len(cols)):
        dfs[i] = df1[start_cols + cols[i]]
        dfs[i] = dfs[i].loc[:,~dfs[i].columns.duplicated()]
        

    df_arr = dfs[0]
    df_scn = dfs[1]
    df_fun = dfs[2]
    df_pth = dfs[3]
    df_tpi = dfs[4]
    df_io = dfs[5]
    df_gt = dfs[6]

    with pd.ExcelWriter(r"test.xlsx") as writer:  
        df1.to_excel(writer, sheet_name='ALL', index=False)

    outfile = dff_tok_path.strip('.csv')

    os.chdir(current_directory)

    sub_dir = os.path.join("outputs//MASTER_" + outfile)
    if not os.path.exists(sub_dir):
        os.mkdir(sub_dir)

    with pd.ExcelWriter(current_directory + "//outputs//MASTER_" + outfile + "//MASTER_" + outfile + ".xlsx") as writer:  
        df_results.to_excel(writer, sheet_name='SUMMARY', index=False)
        df_arr.to_excel(writer, sheet_name='ARR', index=False)
        df_scn.to_excel(writer, sheet_name='SCN', index=False)
        df_pth.to_excel(writer, sheet_name='PTH', index=False)
        df_tpi.to_excel(writer, sheet_name='TPI', index=False)
        df_io.to_excel(writer, sheet_name='IO', index=False)
        df_gt.to_excel(writer, sheet_name='GT', index=False)
        df_fun.to_excel(writer, sheet_name='FUN', index=False)

        for column in df_results:
            column_width = max(df_results[column].astype(str).map(len).max(), len(column))
            col_idx = df_results.columns.get_loc(column)
            writer.sheets['SUMMARY'].set_column(col_idx, col_idx, column_width)

        # Close the Pandas Excel writer and output the Excel file.
        writer.save()
        writer.close() #added to allow time for file lock to be released

    if var1.get() == 1:
        jmp_mod_split_all()
    else:
        print('Analysis Completed Successfully')

def jmp_mod_split_all():
    global output_file

    over_sheet_strip = df_results[['DFF_TokenName', 'ModuleName']] # from the results sheet [each individual delimited token and it's sub module]
    sub_mod_list = df_results['ModuleName'].unique() # all unique submodules
    

    for k in range(0,len(sub_mod_list)): #for each submodule
        split_tokens = [] # reset split tokens
        selected_mod = sub_mod_list[k] # next sub module in list
        if str(selected_mod) == "nan":
            print('skipping nan')
        else:
            for j in range(0,len(over_sheet_strip['ModuleName'])): # each individual delimited token and it's sub module
                if over_sheet_strip['ModuleName'][j] == selected_mod: # if each individual token's submodule matches the selected sub module then and token to a list
                    split_tokens.append(over_sheet_strip['DFF_TokenName'][j]) # have one sub module token list populated here

            # each sub_mod_list needs to know what sheet to go to. should already be done in dfs above, this is the order below
            # arr
            # scn
            # fun
            # pth
            # tpi
            # io
            # gt
            
            if 'ARR' in selected_mod and '_DE' not in selected_mod and '_GT' not in selected_mod:
                sheet = dfs[0]   
            elif 'SCN' in selected_mod and '_DE' not in selected_mod and '_GT' not in selected_mod:
                sheet = dfs[1] 
            elif 'PTH' in selected_mod:
                sheet = dfs[3]
            elif 'TPI' in selected_mod and '_GFX' not in selected_mod:
                sheet = dfs[4] 
            elif 'CLK' in selected_mod:
                sheet = dfs[5]
            elif '_DE' in selected_mod or '_GT' in selected_mod or '_GFX' in selected_mod:
                sheet = dfs[6]
            elif 'FUN' in selected_mod and '_DE' not in selected_mod and '_GT' not in selected_mod:
                sheet = dfs[2]
            else:
                print("can't find token")

            def generate_jmp_report(sheet):

                df_final_list = []
                start_cols=['LOT_OLD','LOT_NEW','OPERATION_OLD','OPERATION_NEW','WAFER_ID','SORT_X','SORT_Y','FBIN_OLD','FBIN_NEW']
                for i in range(0,len(sheet.columns)):
                    if sheet.columns[i] in start_cols:
                        df_final_list.append(sheet.columns[i])
                    else:
                        for j in range(0, len(split_tokens)):
                            if split_tokens[j] in sheet.columns[i]:
                                df_final_list.append(sheet.columns[i])

                df_final = sheet[df_final_list]
                df_final = df_final.loc[:,~df_final.columns.duplicated()]

                output_file = current_directory +"//outputs//jmp_outputs//" + str(selected_mod) + "_DFF_KAPPA.csv"
                print(output_file)
                df_final.to_csv(output_file, encoding='utf-8', index=False)
                run_jmp(output_file,selected_mod)
                sleep(10)

            generate_jmp_report(sheet)
    zip_outputs()

def select_folder():
    global folder_path

    folder_path = fd.askdirectory(
        title='Select DFF KAPPA Folder',
        initialdir='/')

    showinfo(
        title='Selected Folder',
        message=folder_path
    )

def select_file():
    global filename
    filetypes = (
        ('text files', '*.xlsx'),
        ('All files', '*.*')
    )

    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)

    showinfo(
        title='Selected File',
        message=filename
    )
    sub_mod()

def sub_mod():
    global sub_mod_list
    global arr_sheet
    global scn_sheet
    global pth_sheet
    global tpi_sheet
    global io_sheet
    global gt_sheet
    global fun_sheet
    global overview_sheet

    xls = pd.ExcelFile(filename)
    overview_sheet = pd.read_excel(xls, 'SUMMARY')
    arr_sheet = pd.read_excel(xls, modules[0])
    scn_sheet = pd.read_excel(xls, modules[1])
    pth_sheet = pd.read_excel(xls, modules[2])
    tpi_sheet = pd.read_excel(xls, modules[3])
    io_sheet = pd.read_excel(xls, modules[4])
    gt_sheet = pd.read_excel(xls, modules[5])
    fun_sheet = pd.read_excel(xls, modules[6])

    sub_mod_list = overview_sheet['ModuleName'].unique()

    populate_mod_list()

def jmp_mod_split():
    global output_file
    
    selected_mod = variable.get()

    if 'ARR' in variable.get() and '_DE' not in variable.get() and '_GT' not in variable.get():
        sheet = arr_sheet        
    elif 'SCN' in variable.get() and '_DE' not in variable.get() and '_GT' not in variable.get():
        sheet = scn_sheet
    elif 'PTH' in variable.get():
        sheet = pth_sheet
    elif 'TPI' in variable.get() and '_GFX' not in variable.get():
        sheet = tpi_sheet
    elif 'CLK' in variable.get():
        sheet = io_sheet
    elif '_DE' in variable.get() or '_GT' in variable.get() or '_GFX' in variable.get():
        sheet = gt_sheet
    elif 'FUN' in variable.get() and '_DE' not in variable.get() and '_GT' not in variable.get():
        sheet = fun_sheet
    
    # cols = sheet.columns

    split_tokens = []

    over_sheet_strip = overview_sheet[['DFF_TokenName', 'ModuleName']]
    # over_sheet_strip = over_sheet_strip[~over_sheet_strip['DFF_TokenName'].str.contains("^DFF.*_[0-9]*$", na=False) | over_sheet_strip['DFF_TokenName'].str.contains("_0", na=False)]
    # over_sheet_strip['DFF_TokenName'] = over_sheet_strip['DFF_TokenName'].str.replace("_0$", '')
    # over_sheet_strip = over_sheet_strip.reset_index(drop=True)

    for j in range(0,len(over_sheet_strip['ModuleName'])):
        if over_sheet_strip['ModuleName'][j] == selected_mod:
            split_tokens.append(over_sheet_strip['DFF_TokenName'][j])

    sheet_cols = sheet.columns
    final_cols = []

    for i in range(0,len(sheet_cols)):
        for j in split_tokens:
            if j in sheet_cols[i]:
                final_cols.append(sheet_cols[i])


    start_cols=['LOT_OLD','LOT_NEW','OPERATION_OLD','OPERATION_NEW','WAFER_ID','SORT_X','SORT_Y','FBIN_OLD','FBIN_NEW']

    df_final = sheet[start_cols + final_cols]
    df_final = df_final.loc[:,~df_final.columns.duplicated()]

    output_file = current_directory +"//outputs//" + selected_mod + "_DFF_KAPPA.csv"
    print(output_file)
    df_final.to_csv(output_file, encoding='utf-8', index=False)
    run_jmp_ind()
    print('Analysis Completed Successfully')

def run_jmp(output_file = 'null', selected_mod = 'null'):
    # print('before ',os.getcwd())
    os.chdir(current_directory)
    # print('after ',os.getcwd())
    save_path = current_directory +"\outputs\jmp_outputs\\" + str(selected_mod) + "_DFF_KAPPA"
    user_script = current_directory + "\outputs\open_csv.jrp"
    jsl_path = resource_path("Script.jsl")
    reading_file = open(jsl_path, "r")
    
    new_file_content = ""
    for line in reading_file:
        stripped_line = line.strip()
        new_line = stripped_line.replace("txt_path", '"' + output_file + '"')
        new_file_content += new_line +"\n"
    reading_file.close()
    writing_file = open(user_script, "w")
    writing_file.write(new_file_content)
    writing_file.close()

    reading_file = open(user_script, "r")

    new_file_content = ""
    for line in reading_file:
        stripped_line = line.strip()
        new_line = stripped_line.replace("txt_save_path", '"' + save_path)
        new_file_content += new_line +"\n"
    reading_file.close()
    writing_file = open(user_script, "w")
    writing_file.write(new_file_content)
    writing_file.close()
    
    print("Doing JMP");
    print("Running jmp from local computer");
    JMPcaller = r'"'  + resource_path("JMPbackgroundcaller.exe") + '"'  + ' "outputs\open_csv.jrp"'
    print(JMPcaller)
    subprocess.call(JMPcaller,shell = True);

def run_jmp_ind():
    user_script = current_directory + "\outputs\open_csv.jrp"
    jsl_path = resource_path("Script_ind.jsl")
    print('checking',jsl_path)
    reading_file = open(jsl_path, "r")
    
    new_file_content = ""
    for line in reading_file:
        stripped_line = line.strip()
        new_line = stripped_line.replace("txt_path", '"' + output_file + '"')
        new_file_content += new_line +"\n"
    reading_file.close()
    writing_file = open(user_script, "w")
    writing_file.write(new_file_content)
    writing_file.close()
    
    
    os.system('"' + user_script + '"')

def _reset_option_menu(options, index=None):
        '''reset the values in the option menu

        if index is given, set the value of the menu to
        the option at the given index
        '''
        menu = sel_prod["menu"]
        menu.delete(0, "end")
        for string in options:
            menu.add_command(label=string, 
                             command=lambda value=string:
                                  variable.set(value))
        if index is not None:
            variable.set(options[index])

def populate_mod_list():
        '''Switch the option menu to display colors'''
        _reset_option_menu(sub_mod_list, 0)

def zip_outputs():
    dest_para_parent = current_directory + "\\outputs\\MASTER_" + outfile
    dest_para = dest_para_parent + "\\" + outfile

    try:
        os.mkdir(dest_para)
    except:
        print("output files previously generated")

    src_para = current_directory +"\\outputs\\jmp_outputs"
    src_files = os.listdir(src_para)

    print("zipping outputs ")
    for file in src_files:
        if '.csv' in file:
            print("skipping csv files")
        else:
            fn = os.path.join(src_para, file)
            shutil.copy(fn, dest_para)

    shutil.make_archive(dest_para, 'zip', dest_para)
    print("Zipped to: ",outfile)
    print('Analysis Completed Successfully')

def dumpFolder():
    src_para = current_directory + "\\outputs\\jmp_outputs"
    for file in os.listdir(src_para):
        file_path = os.path.join(src_para,file)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))

### Main Root
root = Tk()
root.title('DFF_KAT v1.01')

tab_parent = ttk.Notebook(root)

tab1 = ttk.Frame(tab_parent, padding="60 30 60 50")
tab2 = ttk.Frame(tab_parent, padding="60 50 60 50")

tab_parent.add(tab1, text='Generate Master File')
tab_parent.add(tab2, text='JMP Tables by Module')
tab_parent.grid(sticky=('news'))

def callback(url):
    webbrowser.open_new(url)

link1 = Label(tab1, text="Wiki: https://goto/dffkat", fg="blue", cursor="hand2")
link1.grid(row = 0,column = 0, sticky=W, columnspan = 2)
link1.bind("<Button-1>", lambda e: callback("https://goto/dffkat"))

link2 = Label(tab1, text="IT support contact: idriss.animashaun@intel.com", fg="blue", cursor="hand2")
link2.grid(row = 1,column = 0, sticky=W, columnspan = 2)
link2.bind("<Button-1>", lambda e: callback("https://outlook.com"))

#### Tab 1
open_button = Button(
    tab1,
    text='Select DFF KAPPA Folder',
    command=select_folder,
    bg = 'blue', fg = 'white', font = '-family "SF Espresso Shack" -size 12'
)

open_button.grid(row = 2, column = 0)

button_0 = Button(tab1, text="Generate Master File", height = 1, width = 20, command = master_file, bg = 'green', fg = 'white', font = '-family "SF Espresso Shack" -size 12')
button_0.grid(row = 3, column = 0, rowspan = 2 )

var1 = IntVar(value=1)
Checkbutton(tab1, text="Generate JMP Reports", variable=var1).grid(row=3, column = 1, sticky=W)


#### Tab 2
open_button_1 = Button(
    tab2,
    text='Select MASTER DFF KAPPA File',
    command=select_file,
    bg = 'blue', fg = 'white', font = '-family "SF Espresso Shack" -size 12'
)

open_button_1.grid(row = 2, column = 0)


label_1 = Label(tab2, text = 'Select Module: ', bg  ='black', fg = 'white')
label_1.grid(row = 3, sticky=E)
variable = StringVar(tab2)
variable.set('Select') # default value

sel_prod = OptionMenu(tab2, variable, *sub_modules)
sel_prod.grid(row = 3, column = 1, sticky=W)

button_1 = Button(tab2, text="Open in JMP", height = 1, width = 20, command = jmp_mod_split, bg = 'green', fg = 'white', font = '-family "SF Espresso Shack" -size 12')
button_1.grid(row = 4, column = 0, rowspan = 2 )

### Main loop
root.mainloop()