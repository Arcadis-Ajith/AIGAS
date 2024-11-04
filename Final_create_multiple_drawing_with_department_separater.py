import pandas as pd
import win32com.client
import pythoncom
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import time

def open_directory_dialog(entry):
    directory_path = filedialog.askdirectory(
        title="Select a directory",
        initialdir=os.getcwd()
    )
    if directory_path:
        entry.delete(0, tk.END)
        entry.insert(0, directory_path)

def open_file_dialog(entry_widget, file_type):
    file_paths = filedialog.askopenfilename(
        filetypes=[(f"{file_type} files", f"*.{file_type.lower()}"), ("Excel and CSV files *", "*.xlsx *.csv")],
        title=f"Select {file_type} file"
    )
    if file_paths:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, file_paths)
        return file_paths
    return None

def read_file(file_path, sheet_name=None):
    try:
        if file_path.endswith('.xlsx'):
            return pd.read_excel(file_path, sheet_name=sheet_name)
        elif file_path.endswith('.csv'):
            return pd.read_csv(file_path)
    except Exception as e:
        messagebox.showerror('Error', f"Error reading file: {e}")
        return None

def filter_dataframe(df, column, keyword):
    try:
        return df[df[column].str.contains(keyword, case=False, na=False)]
    except Exception as e:
        messagebox.showerror('Error', f"Error filtering dataframe: {e}")
        return None

def split_long_strings_to_phrases(strings_list, max_length=25):
    result = []
    for string in strings_list:
        if pd.isna(string):
            result.append(['nan'])
            continue
        current_phrase = ""
        phrases = []
        for word in string.split():
            if len(current_phrase) + len(word) + 1 > max_length:
                phrases.append(current_phrase.strip())
                current_phrase = word
            else:
                current_phrase += " " + word
        if current_phrase:
            phrases.append(current_phrase.strip())
        result.append(phrases)
    return result

def on_ok(excel_entry, sheet_combobox, heading_combobox1, heading_combobox2, heading_combobox3, dwg_entry, dwg_entry1, save_entry,save_entry1, save_entry2, save_entry3, save_entry4, root, selected_files):
    excel_file = excel_entry.get()
    selected_sheet = sheet_combobox.get()
    selected_heading1 = heading_combobox1.get()
    selected_heading2 = heading_combobox2.get()
    selected_heading3 = heading_combobox3.get()
    dwg_file = dwg_entry.get()
    dwg_file_tittle_block = dwg_entry1.get()
    save_directory = save_entry.get()
    save_directory1 = save_entry1.get()
    save_directory2 = save_entry2.get() 
    save_directory3 = save_entry3.get() 
    save_directory4 = save_entry4.get()
    
    discipline_name = []
    drawing_number = []
    drawing_title = []

    df = read_file(excel_file, selected_sheet)
    if df is not None:
        filtered_df = filter_dataframe(df, selected_heading2, "dr")
        if filtered_df is not None:
            split_strings = filtered_df[selected_heading3].astype(str).tolist()
            DRAWING_TITLE = split_long_strings_to_phrases(split_strings)
            discipline_name.extend(filtered_df[selected_heading1].astype(str).tolist())
            drawing_number.extend(filtered_df[selected_heading2].astype(str).tolist())
            drawing_title.extend(DRAWING_TITLE)

        if discipline_name and drawing_number and drawing_title and dwg_file and dwg_file_tittle_block and save_directory and save_directory1 and save_directory2 and save_directory3 and save_directory4:
            selected_files.update({
                "discipline_name": discipline_name,
                "drawing_number": drawing_number,
                "drawing_title": drawing_title,
                "dwg_file": dwg_file,
                "dwg_file_tittle_block": dwg_file_tittle_block,
                "save_directory": save_directory,
                "save_directory1": save_directory1,
                "save_directory2": save_directory2,
                "save_directory3" : save_directory3,
                "save_directory4": save_directory4
            })
            root.quit()  # Use quit instead of destroy to keep the widget references valid
            return
        else:
            messagebox.showwarning("Warning", "Please select both files and specify the save directory.")
            return
    root.destroy()

def select_files():
    root = tk.Tk()
    root.title("Select Files and Save Directory")

    selected_files = {}

    tk.Label(root, text="Select MIDP Excel file:").grid(row=0, column=0, padx=10, pady=5)
    excel_entry = tk.Entry(root, width=100)
    excel_entry.grid(row=0, column=1, padx=20, pady=20)
    sheet_combobox = ttk.Combobox(root, width=97)
    tk.Button(root, text="Browse", command=lambda: on_browse(excel_entry, sheet_combobox, heading_combobox1, heading_combobox2, heading_combobox3)).grid(row=0, column=2, padx=10, pady=5)

    tk.Label(root, text="Select sheet Name:").grid(row=1, column=0, padx=10, pady=5)
    sheet_combobox.grid(row=1, column=1, padx=20, pady=20)

    tk.Label(root, text="Select discipline row Heading:").grid(row=2, column=0, padx=10, pady=5)
    heading_combobox1 = ttk.Combobox(root, width=97)
    heading_combobox1.grid(row=2, column=1, padx=20, pady=20)

    tk.Label(root, text="Select drawing number row Heading:").grid(row=3, column=0, padx=10, pady=5)
    heading_combobox2 = ttk.Combobox(root, width=97)
    heading_combobox2.grid(row=3, column=1, padx=20, pady=20)

    tk.Label(root, text="Select drawing Title row Heading:").grid(row=4, column=0, padx=10, pady=5)
    heading_combobox3 = ttk.Combobox(root, width=97)
    heading_combobox3.grid(row=4, column=1, padx=20, pady=20)

    tk.Label(root, text="Select ACAD (standard setting) file:").grid(row=5, column=0, padx=10, pady=5)
    dwg_entry = tk.Entry(root, width=100)
    dwg_entry.grid(row=5, column=1, padx=20, pady=20)
    tk.Button(root, text="Browse", command=lambda: open_file_dialog(dwg_entry, "DWG")).grid(row=5, column=2, padx=10, pady=5)

    tk.Label(root, text="Select Title Block file:").grid(row=6, column=0, padx=10, pady=5)
    dwg_entry1 = tk.Entry(root, width=100)
    dwg_entry1.grid(row=6, column=1, padx=20, pady=20)
    tk.Button(root, text="Browse", command=lambda: open_file_dialog(dwg_entry1, "DWG")).grid(row=6, column=2, padx=10, pady=5)

    tk.Label(root, text="Select civil drawing saving location:").grid(row=7, column=0, padx=10, pady=5)
    save_entry = tk.Entry(root, width=100)
    save_entry.grid(row=7, column=1, padx=20, pady=20)
    tk.Button(root, text="Browse", command=lambda: open_directory_dialog(save_entry)).grid(row=7, column=2, padx=10, pady=5)

    tk.Label(root, text="Select Main Plant drawing saving location:").grid(row=8, column=0, padx=10, pady=5)
    save_entry1 = tk.Entry(root, width=100)
    save_entry1.grid(row=8, column=1, padx=20, pady=20)
    tk.Button(root, text="Browse", command=lambda: open_directory_dialog(save_entry1)).grid(row=8, column=2, padx=10, pady=5)

    tk.Label(root, text="Select P&C drawing saving location:").grid(row=9, column=0, padx=10, pady=5)
    save_entry2 = tk.Entry(root, width=100)
    save_entry2.grid(row=9, column=1, padx=20, pady=20)
    tk.Button(root, text="Browse", command=lambda: open_directory_dialog(save_entry2)).grid(row=9, column=2, padx=10, pady=5)

    tk.Label(root, text="Select RTSD drawing saving location:").grid(row=10, column=0, padx=10, pady=5)
    save_entry3 = tk.Entry(root, width=100)
    save_entry3.grid(row=10, column=1, padx=20, pady=20)
    tk.Button(root, text="Browse", command=lambda: open_directory_dialog(save_entry3)).grid(row=10, column=2, padx=10, pady=5)

    tk.Label(root, text="Select other drawing saving location:").grid(row=11, column=0, padx=10, pady=5)
    save_entry4 = tk.Entry(root, width=100)
    save_entry4.grid(row=11, column=1, padx=20, pady=20)
    tk.Button(root, text="Browse", command=lambda: open_directory_dialog(save_entry4)).grid(row=11, column=2, padx=10, pady=5)

    tk.Button(root, text="OK", command=lambda: on_ok(excel_entry, sheet_combobox, heading_combobox1, heading_combobox2, heading_combobox3, dwg_entry, 
                                                     dwg_entry1, save_entry, save_entry1, save_entry2, save_entry3, save_entry4, root, selected_files)).grid(row=12, column=1, pady=20)

    root.mainloop()
    return selected_files

def on_browse(excel_entry, sheet_combobox, heading_combobox1, heading_combobox2, heading_combobox3):
    file_path = open_file_dialog(excel_entry, "Excel")
    if file_path:
        if file_path.endswith('.xlsx'):
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            sheet_combobox['values'] = sheet_names

            def on_sheet_select(event):
                selected_sheet = sheet_combobox.get()
                df = pd.read_excel(file_path, sheet_name=selected_sheet)
                headings = df.columns.tolist()
                heading_combobox1['values'] = headings
                heading_combobox2['values'] = headings
                heading_combobox3['values'] = headings

            sheet_combobox.bind("<<ComboboxSelected>>", on_sheet_select)
        
        elif file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
            headings = df.columns.tolist()
            heading_combobox1['values'] = headings
            heading_combobox2['values'] = headings
            heading_combobox3['values'] = headings

def process_file_path(file_path, dwg_files):
    if file_path:
        file_name = os.path.basename(file_path)
        file_location = os.path.dirname(file_path)
        dwg_files.append((file_name, file_path))

def create_insertion_point(info):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, info)

def get_relative_path(base_path, target_path):
    return os.path.relpath(target_path, base_path)

# Convert the absolute path to a relative path
current_working_directory = os.getcwd()
base_path = r"C:\Users\mahadeva1972\DC"  # Explicitly set the base path

def change_attribute_value(doc, block_name, tags, new_values, drawing_titles):
    if len(drawing_titles) == 4:
        tags_to_edit = ["TITLE_TEXT_LINE_01", "TITLE_TEXT_LINE_02", "TITLE_TEXT_LINE_03", "TITLE_TEXT_LINE_04"]
    elif len(drawing_titles) == 3:
        tags_to_edit = ["TITLE_TEXT_LINE_02", "TITLE_TEXT_LINE_03", "TITLE_TEXT_LINE_04"]
    elif len(drawing_titles) == 2:
        tags_to_edit = ["TITLE_TEXT_LINE_02", "TITLE_TEXT_LINE_03"]
    else:
        tags_to_edit = ["TITLE_TEXT_LINE_02"]
    time.sleep(1)
    for entity in doc.PaperSpace:
        if entity.EntityName == 'AcDbBlockReference' and entity.Name == block_name:
            for attribute in entity.GetAttributes():
                try:
                    if attribute.TagString == tags:
                        attribute.TextString = str(new_values)
                    elif attribute.TagString in tags_to_edit:
                        index = tags_to_edit.index(attribute.TagString)
                        if 0 <= index < len(drawing_titles):
                            attribute.TextString = str(drawing_titles[index])
                        else:
                            print(f"Invalid index: {index}")
                    attribute.Update()
                except Exception as e:
                    print(f"COM error: {e}")
                    time.sleep(1)  # Add a delay before retrying
                    attribute.Update()  # Retry the update

def automate_autocad(dwg_file,dwg_file_tittle_block,drawing_numbers,drawing_titles,layer_name,discipline_names,civil_save_directory,Main_plant_save_directory1,P_C_save_directory2,RTSD_save_directory3,other_save_directory4):
    acad = win32com.client.Dispatch("AutoCAD.Application")
    acad.Visible = True
    insertion_point = (841, 0, 0)
    Xref_insertionpoint = (0, 0, 0)
    block_name = "DESCRIPTION & REVISION"
    attName = "ARCADIS_DRAWING_NUMBER"
    tStart = time.time()

    for i in range(len(drawing_numbers)):
        drawing_number = drawing_numbers[i]
        drawing_title = drawing_titles[i]
        try:
            doc = acad.Documents.open(dwg_file)
            time.sleep(10)
        except Exception as e:
            print(f"Error opening document: {e}")
            continue

        paper_space = doc.PaperSpace
        doc = acad.ActiveDocument
        layout = doc.Layouts.Item("Layout1")
        doc.ActiveLayout = layout
        try:
            for file_name1, file_path1 in dwg_file_tittle_block:
                relative_xref_path = get_relative_path(current_working_directory, file_path1)
                try:
                    xref = paper_space.AttachExternalReference(relative_xref_path, file_name1, create_insertion_point(Xref_insertionpoint), 1, 1, 1, 0, True, False)
                    xref.Layer = layer_name
                    file_name1 = file_name1.replace('.dwg', '')
                    #doc.SendCommand(f'-XREF\nt\n{file_name1}\nr\n')
                   #xref.Path = "C:\\Users\\mahadeva1972\\DC\\ACCDocs\\Arcadis ACC EU\\AGB-30238794-Aigas\\Project Files\\05-Resources\\A_Templates\\SSE\\XXXXXXXX-ARC-SUB-ZZ -DR-EE-0000-TITLE BLOCK A1.dwg"  # Replace with the new path
                    #xref.Reload()
                except Exception as e:
                    print(f"Failed to add Xref from {file_path1}: {e}")
        except Exception as e:
            print(f"Error attaching Xref1: {e}")
        paper_space.InsertBlock(create_insertion_point(insertion_point), block_name, 1, 1, 1, 0)
        drawing_number = drawing_number.replace('.', '')
        layout.Name = drawing_number
        print(drawing_number, end="")
        print(f"-{discipline_names[i]}")
        change_attribute_value(doc, block_name, attName, drawing_number,drawing_title)
        if discipline_names[i] == "civil":
            doc.SaveAs(os.path.join(civil_save_directory, f"{drawing_number}.dwg"))
        elif discipline_names[i] == "Main Plant":
            doc.SaveAs(os.path.join(Main_plant_save_directory1, f"{drawing_number}.dwg"))
        elif discipline_names[i] == "P&C":
            doc.SaveAs(os.path.join(P_C_save_directory2, f"{drawing_number}.dwg"))
        elif discipline_names[i] == "RTSD":
            doc.SaveAs(os.path.join(RTSD_save_directory3, f"{drawing_number}.dwg"))
        else:
            doc.SaveAs(os.path.join(other_save_directory4, f"{drawing_number}.dwg"))
        time.sleep(10) # Increase sleep time before closing the document
        doc.Close()
    tEnd = time.time()
    FullTime = tEnd - tStart
    Total_drawing = len(drawing_number)
    print(f"I have created {Total_drawing} drawings within {FullTime} seconds")

if __name__ == "__main__":
    selected_files = select_files()
    if selected_files:
        discipline_names = selected_files['discipline_name']
        drawing_numbers = selected_files['drawing_number']
        drawing_titles = selected_files['drawing_title']
        dwg_file = selected_files['dwg_file']
        dwg_file_tittle_block =  selected_files['dwg_file_tittle_block']
        civil_save_directory = selected_files['save_directory']
        Main_plant_save_directory1 = selected_files['save_directory1']
        P_C_save_directory2 = selected_files['save_directory2']
        RTSD_save_directory3 = selected_files['save_directory3']
        other_save_directory4 = selected_files['save_directory4']
        template = []
        layer_name = "0"
        process_file_path(dwg_file_tittle_block, template)
        automate_autocad(dwg_file,template,drawing_numbers,drawing_titles,layer_name,discipline_names,civil_save_directory,Main_plant_save_directory1,P_C_save_directory2,RTSD_save_directory3,other_save_directory4)
        '''print(f"Selected DWG File: {dwg_file}")
        print(f"Selected DWG Title Block File: {dwg_file_tittle_block}")
        print(f"Selected save_directory File: {save_directory}")'''
