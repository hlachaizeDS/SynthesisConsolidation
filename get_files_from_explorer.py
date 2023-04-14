import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
import ttkwidgets
import os
from pathlib import Path

initial_path=Path("C:\\Users\\HL\\DNA Script\\Thomas YBERT - SYNTHESIS OPERATIONS\\S.3 - P4\\Quartets\\2110")

tmp_data="path_lists\\tmp.data"
tmp_last_data="path_lists\\tmp_last.data"

def getFileList(initial_path,contains):

    contains_list=contains.split(',')
    files = []
    # r=root, d=directories, f = files
    for r, d, f in os.walk(initial_path):
        for file in f:
            if all(cont in file for cont in contains_list) and file[0] != "~" and 'All' not in r:
                #We keep only the name of the folders right after the initial dir, not to have too many
                #Except if we're already on the root
                path_without_initial_and_splitted=r.replace(str(initial_path),'').split("\\")
                if len(path_without_initial_and_splitted)>1:
                    files.append([file,r + "\\" + file , path_without_initial_and_splitted[1]])
                else:
                    files.append([file, r + "\\" + file, path_without_initial_and_splitted[0]])

    return files


def chose_files_from_explorer(initial_path,initial_str):

    root=tk.Tk()

    #initial path label
    path_label_string = StringVar()
    path_label_string.set(initial_path)
    path_label=tk.Label(root,textvariable=path_label_string)
    path_label.grid(row=1,column=1,columnspan=3,pady=2)

    path_change_button = tk.Button(text="Change Path",
                                   command=lambda: change_folder_callback(tree_view, path_label_string,restriction_string))
    path_change_button.grid(row=2,column=1,columnspan=3,pady=2)

    open_path_file_button = tk.Button(text="Load paths",
                                   command=lambda: check_paths(tree_view,path_label_string,restriction_string,1))
    open_path_file_button.grid(row=3,column=1,pady=2,sticky="e")

    open_last_button = tk.Button(text="Load last",
                                      command=lambda: check_paths(tree_view,path_label_string,restriction_string,0))
    open_last_button.grid(row=3, column=2,pady=2)

    save_selection_button = tk.Button(text="Save selection",
                                 command=lambda: save_selection(tree_view))
    save_selection_button.grid(row=3, column=3,pady=2,sticky="w")

    #Name restriction
    restriction_string = StringVar()
    restriction_string.set(initial_str)
    restriction_label = tk.Entry(root, textvariable=restriction_string)
    restriction_label.grid(row=4,column=1,columnspan=3,pady=2)

    restriction_change_button = tk.Button(text="Change Restriction Strings",
                                   command=lambda: change_restriction_callback(tree_view, path_label_string,restriction_string))
    restriction_change_button.grid(row=5,column=1,columnspan=3,pady=2)

    # Tree view
    scroll = ttk.Scrollbar(root, orient=tk.VERTICAL)
    tree_view = ttkwidgets.CheckboxTreeview(root, yscrollcommand=scroll.set)
    scroll.config(command=tree_view.yview)
    # columns
    tree_view["columns"] = ("Path")  # There's one extra column at the beginning for checkboxes/images
    tree_view["displaycolumns"] = []

    tree_view.column("#0", width=700, minwidth=700, stretch=tk.YES)
    tree_view.column("Path", width=270, minwidth=270, stretch=tk.YES)

    tree_view.heading("#0", text="File name", anchor=tk.W)
    tree_view.heading("Path", text="Path", anchor=tk.W)

    tree_view.grid(row=7,column=1,columnspan=3,pady=2)
    scroll.grid(row=7,column=4, sticky="ns")

    files=getFileList(initial_path,restriction_string.get())

    fill_tree_view(tree_view,files,"")

    #Select Button
    select_button = tk.Button(text="Select",
                                          command=lambda:root.quit())
    select_button.grid(row=6,column=1,columnspan=3,pady=2)

    root.mainloop()

    checked_paths = get_checked(tree_view)
    unchecked_paths = get_unchecked(tree_view)
    tmp_paths = read_list(tmp_data)
    paths_to_save = list(set(tmp_paths + checked_paths) - set(unchecked_paths))
    store_list(tmp_data, paths_to_save)
    store_list(tmp_last_data, paths_to_save)

    return paths_to_save

def fill_tree_view(tree_view,files,already_checked):
    current_folder = None
    for f in files:
        if f[2] != current_folder:
            current_folder = f[2]
            tree_folder = tree_view.insert("", "end", text=current_folder, values=(''))
        if f[1]in already_checked:
            tree_view.insert(tree_folder, "end", text=f[0], values=(f[1],), tags="checked") #need the comma in values apparently
        else:
            tree_view.insert(tree_folder, "end", text=f[0], values=(f[1],)) #need the comma in values apparently

def check_paths(tree_view,path_label_string,restriction_string,fromFile):
    if fromFile:
        path_file=filedialog.askopenfilename(title="Choose path file",initialdir="path_lists")
    else:
        path_file=tmp_last_data
    paths_to_check=read_list(path_file)
    tree_view.delete(*tree_view.get_children())
    files = getFileList(path_label_string.get(),restriction_string.get())
    fill_tree_view(tree_view, files,paths_to_check)
    store_list(tmp_data,paths_to_check)
    store_list(tmp_last_data, paths_to_check)

def save_selection(tree_view):
    checked_paths = get_checked(tree_view)
    unchecked_paths = get_unchecked(tree_view)
    tmp_paths = read_list(tmp_data)
    paths_to_save = list(set(tmp_paths + checked_paths)-set(unchecked_paths))
    file_to_save=filedialog.asksaveasfilename(title="Chose path file to save to",initialdir="path_lists",defaultextension=".data")
    store_list(file_to_save,paths_to_save)
    store_list(tmp_data,paths_to_save)
    store_list(tmp_last_data, paths_to_save)


def change_folder_callback(tree_view,path_label_string,restriction_string):
    checked_paths = get_checked(tree_view)
    unchecked_paths = get_unchecked(tree_view)
    tmp_paths = read_list(tmp_data)
    paths_to_save = list(set(tmp_paths + checked_paths) - set(unchecked_paths))
    store_list(tmp_data,paths_to_save)
    store_list(tmp_last_data,paths_to_save)
    store_list("path_lists\\tmp.data",paths_to_save)
    tree_view.delete(*tree_view.get_children())
    new_path=filedialog.askdirectory(title="Choose directory",initialdir=path_label_string.get()).replace("/","\\")
    files=getFileList(new_path,restriction_string.get())
    fill_tree_view(tree_view,files,paths_to_save)
    path_label_string.set(new_path)

def change_restriction_callback(tree_view,path_label_string,restriction_string):
    checked_paths = get_checked(tree_view)
    unchecked_paths = get_unchecked(tree_view)
    tmp_paths = read_list(tmp_data)
    paths_to_save = list(set(tmp_paths + checked_paths) - set(unchecked_paths))
    store_list(tmp_data,paths_to_save)
    store_list(tmp_last_data,paths_to_save)
    tree_view.delete(*tree_view.get_children())
    files=getFileList(path_label_string.get(),restriction_string.get())
    fill_tree_view(tree_view,files,paths_to_save)

def get_checked(tree):
    checked = []
    def rec_get_checked(item):
        if tree.tag_has('checked', item):
            item_value=tree.item(item,"values")
            if item_value!='':
                checked.append(item_value[0])
        for ch in tree.get_children(item):
            rec_get_checked(ch)

    rec_get_checked('')
    return checked

def get_unchecked(tree):
    unchecked = []
    def rec_get_unchecked(item):
        if tree.tag_has('checked', item) != True:
            item_value=tree.item(item,"values")
            if item_value!='':
                unchecked.append(item_value[0])
        for ch in tree.get_children(item):
            rec_get_unchecked(ch)

    rec_get_unchecked('')
    return unchecked

def store_list(data_file_path,file_paths):
    with open(data_file_path, 'w') as filehandle:
        filehandle.writelines("%s\n" % path for path in file_paths)

def read_list(data_file_path):

    file_paths=[]
    with open(data_file_path, 'r') as filehandle:
        filecontents = filehandle.readlines()

    for line in filecontents:
        # remove linebreak which is the last character of the string
        current_place = line[:-1]
        # add item to the list
        file_paths.append(current_place)

    return file_paths


if __name__ == '__main__':
    file_paths=chose_files_from_explorer(initial_path,"_P4_")
    print(file_paths)
    store_list("tmp.data",file_paths)
    paths=read_list("tmp.data")
    print(paths)