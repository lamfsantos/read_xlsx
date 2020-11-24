import os
import sys
import platform
import pandas as pd
import tkinter as tk
from tkinter.filedialog import askopenfilename

def create_view():
    window = tk.Tk()
    window.wm_title("Read xlsx")
    window.geometry('300x120')

    txt_label = tk.Label(text='Insert your xlsx file:')
    txt_label.pack()

    txt_space = tk.Label(text='')
    txt_space.pack()

    btn_search = tk.Button(text='Search', padx=20, command=process_xlsx)
    btn_search.pack()

    return window

def done():
    txt_label = tk.Label(text="")
    txt_label.pack()
    txt_label = tk.Label(text="Done! Check your output folder.")
    txt_label.pack()

def read_xlsx_document(path: str):

    df = pd.read_excel(path)
    df = df.drop([0,0])

    header_row = 0
    df.columns = df.iloc[header_row]
    df = df.drop([1,1])

    df = df[['Keyword', 'Competition', 'Competition (indexed value)']]

    return df

def sort_xlsx(df):
    df = df.sort_values(by=['Competition', 'Competition (indexed value)'] )
    return df

def filter_xlsx_by_word(df):
    filtered_keywords = []
    list_keywords = df['Keyword'].values.tolist()

    for keyword in list_keywords:
        if len(filtered_keywords) == 0:
            filtered_keywords.append(keyword)
        else:
            should_add = True
            for filtered_keyword in filtered_keywords:
                if filtered_keyword in keyword or keyword in filtered_keyword:
                    should_add = False

            if should_add:
                filtered_keywords.append(keyword)

    return filtered_keywords

def get_file():
    #tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    try:
        filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file
        return(filename)
    except Exception as e:
        raise e

def output_xlsx(pd, path, filtered_keywords):
    #print(type(pd))
    for index, row in pd.iterrows():
        should_delete = False
        for filtered_keyword in filtered_keywords:
            if row['Keyword'] == filtered_keyword:
               should_delete = True

        if not should_delete: 
            pd = pd.drop(index)

    pd = pd.to_excel(path)

def output_path():
    dir_path = os.path.dirname(os.path.realpath(__file__))
    return dir_path+key_per_os()+"output_files"

def split_path(path):
    path_array = path.split("/")
    key = key_per_os()
    return key + path_array[len(path_array)-1]

def key_per_os():
    key = ''
    if platform.system() == 'Linux':
        key = "/"
    else:
        key = "\\"

    return key

def process_xlsx():
    input_file = get_file()
    #print(input_file)
    pd = read_xlsx_document(input_file)
    pd = sort_xlsx(pd)
    output_folder_path = output_path()+split_path(input_file)
    filtered_keywords = filter_xlsx_by_word(pd)
    output_xlsx(pd, output_folder_path, filtered_keywords)
    done()

if __name__ == '__main__':
    window = create_view()
    window.mainloop()
