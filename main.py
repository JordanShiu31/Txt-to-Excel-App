import openpyxl as op
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import shutil
#-----------------------------------------------------#
# ---------- Create the application window ---------- #
#-----------------------------------------------------#

window = tk.Tk()
window.title("Transferring Txt data into Excel")
window.geometry('1080x720')
window.minsize(1080,720)
selected_worksheet_names = []
scenario_state = []
scenario_years = []
scenario_period = []
event_enabled = True

#-----------------------------------------------------#
# ---------- Functions to operate the GUI  ---------- #
#-----------------------------------------------------#
#-------------------------------------------------------------------------------------------#
## Functions for TEXTBOX Widget

def input_to_list(input_word):
    new_list = []
    temp_list = input_word.split(",")
    for word in temp_list:
        new_list.append(word.strip().upper())
    return new_list

def update_txt_files(curr_listbox, functionality):
    replacing_dict = {}
    root_directory_path = txt_root_directory_entry.get().upper()
    common_file_name = txt_result_file_names_entry.get().upper()
    existing_word = txt_exisitng_name_entry.get().upper()
    replacing_word = txt_replacing_name_entry.get().upper()
    destination_path = txt_final_folder_path_entry.get().upper()

    existing_word = input_to_list(existing_word)
    replacing_word = input_to_list(replacing_word)

    # adds list of existing/replacing words to a dict
    if len(existing_word) == len(replacing_word) and functionality == "UPDATE TXT":
        for index, key in enumerate(existing_word):
            replacing_dict.update({key:replacing_word[index]})
    elif functionality == "SHOW TXT":
        pass
    else:
        messagebox.showerror("Error", f"The list size: {existing_word} is different to {replacing_word}")
        return
    
    count = 0
    no_file_updated = 0

    if os.path.exists(root_directory_path):
        list_of_found_txt = read_files_from_folders(root_directory_path, common_file_name)
        if len(list_of_found_txt[0]) == 0:
            messagebox.showerror("Error", f"No files ending with {common_file_name} was found in {root_directory_path}!\nPlease check the entered path name")
        elif functionality == "SHOW TXT":
            # if not os.path.exists(destination_path):
            #     messagebox.showerror("Error", f"The path: {destination_path} does not exist!\nPlease enter an exisiting folder")
            # else:
            txt_show_updated_button.config(state=tk.NORMAL)
            update_listbox(list_of_found_txt[1], curr_listbox)
        elif functionality == "UPDATE TXT":
            for filename in list_of_found_txt[1]:
                for key, value in replacing_dict.items():
                    if key in filename:
                        count+=1
            if count < 1:
                messagebox.showerror("Error", f"The word(s): {existing_word} cannot be found in TXT filest!\nPlease enter an exisiting word")
            else:
                txt_delete_new_files_button.config(state=tk.NORMAL)
                for curr_path_name, curr_file_name in zip(list_of_found_txt[0],list_of_found_txt[1]):
                    new_file_name = curr_file_name.upper()
                    for key, value in replacing_dict.items(): 
                        if key in curr_file_name:
                            new_file_name = new_file_name.replace(key, value)
                    final_destination_path = os.path.join(destination_path, new_file_name)
                    shutil.copy(curr_path_name, final_destination_path)
                    no_file_updated+=1
                new_file_list = read_files_from_folders(destination_path, common_file_name)
                update_listbox(new_file_list[2], curr_listbox)
                messagebox.showinfo("Success!",f"{no_file_updated} has been updated")
    else:
        messagebox.showerror("Error", "Incorrect Folder Directory!\nPlease check the entered path name")

def browse(cur_stringvar, is_dir=True):
    print(cur_stringvar.get())
    if is_dir:
        file_path = filedialog.askdirectory()
    else:
        file_path= filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls;*.xlsx")])
    if file_path:
        print("it should work and replace")
        global event_enabled
        event_enabled = False
        cur_stringvar.set(file_path)


## Removes white spaces and seperates each line to a list
def convert_string_to_list(curr_list, curr_text, button, button_name):
    data = curr_text.get("1.0","end")
    curr_list.append(data)
    final_list = data.replace(" ", "").strip().split("\n")
    if(len(final_list) > 0):
        scenario_param_button_text_update(final_list, button, button_name)

# Updates button for TextBox and assigns contents to the respective lists
def scenario_param_button_text_update(textbox_list_items, button, button_var_name):
    global scenario_years
    global scenario_state
    selected_text = ""
    if button_var_name == "scenario_years_button":
        scenario_years = textbox_list_items
        if len(textbox_list_items) == 1 and textbox_list_items[0] == "":
            selected_text = f"Please enter a year(s)"
        else:
            selected_text = f"{len(textbox_list_items)} years selected"
    elif button_var_name == "scenario_state_button":
        scenario_state = textbox_list_items
        if len(textbox_list_items) == 1 and textbox_list_items[0] == "":
            selected_text = f"Please enter a state(s)"
        else:
            selected_text = f"{len(textbox_list_items)} states selected"
    else:
        pass
    button.config(text=selected_text)

#-------------------------------------------------------------------------------------------#

#-------------------------------------------------------------------------------------------#
def delete_selected_listbox_items(curr_listbox):
    selected_indices = curr_listbox.curselection() # 
    # assumes windows will always read in order (might be diff for other systems)
    for index, file in enumerate(os.listdir(txt_final_folder_path_string.get())):
        if index in selected_indices:
            file_path = os.path.join(txt_final_folder_path_string.get(), file)
            os.remove(file_path)
    for index in selected_indices[::-1]:
        curr_listbox.delete(index)

## Returns a list of selected tab names
def get_selected_listbox_items(curr_listbox):
    selected_indices = curr_listbox.curselection() # 
    selected_listbox_items = [curr_listbox.get(index) for index in selected_indices]
    return selected_listbox_items

## Updates the number of periods selected and the button text
def assign_scenario_period_to_list(curr_listbox, button):
    global scenario_period
    scenario_period = get_selected_listbox_items(curr_listbox)
    selected_text = f"{len(scenario_period)} periods selected"
    button.config(text=selected_text)
#-------------------------------------------------------------------------------------------#

#-------------------------------------------------------------------------------------------#
## Functions for EXCEL LISTBOX Widget
## Updates the number of worksheets selected and the button text
def get_selected_worksheet_no(curr_listbox, button):
    selected_worksheet_names = get_selected_listbox_items(curr_listbox)
    selected_text = f"{len(selected_worksheet_names)} worksheets selected"
    button.config(text=selected_text)
    update_excel_button.config(state=tk.NORMAL) # unlocks the update button

## Adds existing worksheet tabs into the listbox    
def update_listbox(list_of_items, curr_listbox):
    curr_listbox.delete(0, tk.END)
    for item in list_of_items:
        curr_listbox.insert(tk.END, item)

## deletes all data in the listbox
# attempts to open user input excel file and will display errors if excel file is not valid
# adds all exisitng excel tabs (worksheets) into the list
def get_worksheet_names(excel_stringvar, listbox):
    # converts to a string
    file_path = excel_stringvar.get()
    # excel_worksheet_listbox.delete(0, tk.END)
    try:
        workbook = op.load_workbook(file_path)
        update_listbox(workbook.sheetnames, listbox)
        workbook.close()
        error_msg.set(value="Valid Spreadsheet")
        select_worksheet_button.config(state=tk.NORMAL) # NEED TO ADD IF PARAMS ARE SET AS WELL
    except op.utils.exceptions.InvalidFileException:
        error_msg.set(value="Error: Invalid path, please enter a correct excel path")
    except FileNotFoundError:
        error_msg.set(value="Error: Excel file not found")
#-------------------------------------------------------------------------------------------#

## Returns a list of keywords that matches the user input and user desired worksheet names
def find_worksheet_key_words(worksheet_name):
    desired_words = []
    caps_worksheet_name = worksheet_name.upper()
    for state in scenario_state:
        if state.upper() in caps_worksheet_name:
            for period in scenario_period:
                if period.upper() in caps_worksheet_name:
                    for year in scenario_years:
                        if year in caps_worksheet_name:
                            desired_words.append(state.upper())
                            desired_words.append(period)
                            desired_words.append(year)
    return desired_words

# Returns a list containing the FULL txt file paths and the Txt file names
def read_files_from_folders(root_directory, common_ending_file_name, check_same_dir=True):
    list_raw_data_path = []
    list_raw_file_name = []
    list_raw_dest_files = []
    list_raw_path_and_name = []

    try:
        for file in os.listdir(txt_final_folder_path_string.get()):
            list_raw_dest_files.append(file.upper())
    except OSError:
        messagebox.showerror("Error", f"The path: {txt_final_folder_path_string.get()} does not exist!\nPlease enter an exisiting folder")
        return
    for root, _, files in os.walk(root_directory):
        if os.path.basename(root.upper()) == os.path.basename(txt_final_folder_path_string.get().upper()) and check_same_dir:
            pass
        else:
            for file in files:
                cap_file_name = common_ending_file_name.upper()
                _file = file.upper()
                if _file.endswith(cap_file_name):
                    caps_file_path = os.path.join(root.upper(), _file)
                    list_raw_file_name.append(_file)       
                    list_raw_data_path.append(caps_file_path)
    list_raw_path_and_name.append(list_raw_data_path)
    list_raw_path_and_name.append(list_raw_file_name)
    list_raw_path_and_name.append(list_raw_dest_files)
    return list_raw_path_and_name

## Returns a list of the full txt file paths which contain the desired keywords
def find_txt_file(keywords_list, raw_data_files_list): 
    for index, path_name in enumerate(raw_data_files_list[1]):
        counter = 0
        for word in keywords_list:
            if word in path_name:
                counter+=1
        if counter == len(keywords_list):
            return raw_data_files_list[0][index]
    return False        

## Opens the selected full txt file path and reads all lines, clears all exisitng excel contents and pasts new contents
def open_txt_file(desired_path_file, delimiter, worksheet_active):
    with open(desired_path_file, 'r') as file:
        lines = file.readlines()
        clear_worksheet_contents(worksheet_active)
        paste_contents(lines, delimiter, worksheet_active)
        
## clears all exsisting data inside the active worksheet
def clear_worksheet_contents(worksheet_active):
    for row in worksheet_active.iter_rows():
        for cell in row:
            cell.value = None

## pastes each line of txt data into the active worksheet and attempts to paste them as int value type
def paste_contents(lines, delimiter, worksheet_active):
    start_row = 1
    start_column = 1
    for row_idx, line in enumerate(lines, start=start_row):
        values = line.strip().split(delimiter)  # Replace ',' with your desired delimiter
        for col_idx, value in enumerate(values, start=start_column):
            worksheet_active.cell(row=row_idx, column=col_idx, value=value)
            if value.isdigit():
                worksheet_active.cell(row=row_idx, column=col_idx, value=int(value))
            else:
                try:
                    worksheet_active.cell(row=row_idx, column=col_idx, value=float(value))
                except ValueError:
                    worksheet_active.cell(row=row_idx, column=col_idx, value=value)

## the MAin function which does the magic when the update excel button is pressed
def update_func(curr_listbox):
    excel_path = excel_string.get()
    root_directory_path = root_directory_entry.get()
    caps_root_directory_path = root_directory_path.upper()
    common_ending_file_name = txt_result_file_names_string.get()
    delimiter = delimiter_entry.get()
    selected_worksheet_names = get_selected_listbox_items(curr_listbox)
    workbook = op.load_workbook(excel_path)
    
    unselected_worksheets = []
    pasted_contents_worksheets = []
    
    if os.path.exists(caps_root_directory_path):
        raw_data_files_list = read_files_from_folders(caps_root_directory_path, common_ending_file_name, False)
        # if the root fodler is found but no result files could be found it could be an error
        if len(raw_data_files_list[0]) == 0:
            show_error_message_box(raw_data_files_list[0], root_directory_path, common_ending_file_name)
        else: # start to algorithm
            for worksheet_name in selected_worksheet_names:
                worksheet_active = workbook[worksheet_name]
                keywords_list = find_worksheet_key_words(worksheet_name)
                if len(keywords_list) >= 3:
                    desired_path_file = find_txt_file(keywords_list, raw_data_files_list)             
                    if desired_path_file == False:
                        messagebox.showerror("Error", f"raw data file could not be found! for:\n {worksheet_name}")
                        unselected_worksheets.append(worksheet_name)
                    else:
                        open_txt_file(desired_path_file, delimiter, worksheet_active)
                        pasted_contents_worksheets.append(worksheet_name)
                else:
                    unselected_worksheets.append(worksheet_name)
            workbook.save(excel_path)
    else:
        messagebox.showerror("Error", "Incorrect Folder Directory!\nPlease check the entered path name")
    workbook.close()
    messagebox.showinfo("Complete", f"Data has been pasted in Excel for scenarios:\n{pasted_contents_worksheets}\nHowever, scenrios:\n{unselected_worksheets}\nwere not pasted due to user scenario parameters")

## Pop up errorbox for no valid files found in folder
def show_error_message_box(raw_data_files_list, root_directory_path, common_ending_file_name):
    message = (
        f"There were {len(raw_data_files_list)} found files found in folder "
        f"with keyworks :\n{common_ending_file_name}\n\n"
        f"Please double check the 'Raw Results File Name' or there may be "
        f"no results at all inside:\n{root_directory_path}"
    )
    messagebox.showerror("Error", message)
    
def on_tab_selected(event):
    selected_tab = tab_control.index(tab_control.select())
    
#-------------------------------------------------------------------------------------------#
#-------------------------------------------------#
# ---------- Creating all the widgets  ---------- #
#-------------------------------------------------#
tab_control = ttk.Notebook(window)

# 1. Creates a frame inside window to allow some padding and to help with widget placement
txt_frame = tk.Frame(tab_control)
txt_frame.pack(expand=True, fill="both", padx=10, pady=10) 
tab_control.add(txt_frame, text="Txt File Renaming")

frame = tk.Frame(tab_control)
frame.pack(expand=True, fill="both", padx=10, pady=10) 
tab_control.add(frame, text="Txt to Excel")

tab_control.bind("<<NotebookTabChanged>>", on_tab_selected)
tab_control.pack(expand=1, fill="both")
# ---------- tabs 1 Widgets  ---------- #
# 2. Creates subframes inside fame above to allow grouping of other widgets
txt_input_frame = tk.Frame(txt_frame)
txt_display_frame = tk.Frame(txt_frame)

txt_path_files_text_frame = tk.Frame(txt_input_frame)
txt_path_files_entry_frame = tk.Frame(txt_input_frame)
txt_path_files_entry_button_frame = tk.Frame(txt_input_frame)
txt_display_existing_frame = tk.Frame(txt_display_frame)
txt_display_updated_frame = tk.Frame(txt_display_frame)

# Label Widgets
txt_root_directory_label = ttk.Label(txt_path_files_text_frame, text="Full Folder Path")
txt_result_file_names_label = ttk.Label(txt_path_files_text_frame, text="Common Txt File name")
txt_exisitng_name_label = ttk.Label(txt_path_files_text_frame, text="Desired word to replace")
txt_replacing_name_label = ttk.Label(txt_path_files_text_frame, text="New word to replace old word")
txt_final_folder_path_label = ttk.Label(txt_path_files_text_frame, text="Destination Folder")

txt_root_directory_string = tk.StringVar(value='Please browse or enter the full folder path file cotaining all result files with no "" marks')
txt_result_file_names_string = tk.StringVar(value="Please enter a common results file name. eg.Vehicle Travel Time Results.att")
txt_exisitng_name_string = tk.StringVar(value="Please enter an exisitng word(s) you want to replace in the txt file names")
txt_replacing_name_string = tk.StringVar(value="Please enter a desired word(s) you want to update in the txt file names")
txt_final_folder_path_string = tk.StringVar(value='Please browse or enter a destination folder path without the "" marks')

# Listbox Widgets
txt_existing_error_msg = tk.StringVar(value="Click below to verify exisitng files")
txt_updated_error_msg = tk.StringVar(value="Click below to verify updated files")
txt_existing_file_label = ttk.Label(txt_display_existing_frame, textvariable=txt_existing_error_msg)
txt_updated_file_label = ttk.Label(txt_display_updated_frame, textvariable=txt_updated_error_msg)
txt_exisitng_path_name_listbox = tk.Listbox(txt_display_existing_frame)
txt_updated_path_name_listbox = tk.Listbox(txt_display_updated_frame, selectmode=tk.EXTENDED)

# Entry Widgets
entry_width = 100
txt_root_directory_entry = ttk.Entry(txt_path_files_entry_frame, textvariable=txt_root_directory_string, width=entry_width)
txt_result_file_names_entry = ttk.Entry(txt_path_files_entry_frame, textvariable=txt_result_file_names_string, width=entry_width)
txt_exisitng_name_entry = ttk.Entry(txt_path_files_entry_frame, textvariable=txt_exisitng_name_string, width=entry_width)
txt_replacing_name_entry = ttk.Entry(txt_path_files_entry_frame, textvariable=txt_replacing_name_string, width=entry_width)
txt_final_folder_path_entry = ttk.Entry(txt_path_files_entry_frame, textvariable=txt_final_folder_path_string, width=entry_width)

# Button Widgets
txt_root_dir_browse_button = tk.Button(txt_path_files_entry_button_frame, text="Browse Folder", width=20, command=lambda:browse(txt_root_directory_string))
txt_dest_dir_browse_button = tk.Button(txt_path_files_entry_button_frame, text="Browse Destination", width=20, command=lambda:browse(txt_final_folder_path_string))
txt_show_existing_button = ttk.Button(txt_display_existing_frame, text="Show exisitng txt file names", command=lambda:update_txt_files(txt_exisitng_path_name_listbox, "SHOW TXT"))
txt_show_updated_button = ttk.Button(txt_display_updated_frame, text="Show updated txt file names", command=lambda:update_txt_files(txt_updated_path_name_listbox, "UPDATE TXT"), state=tk.DISABLED)
txt_delete_new_files_button = ttk.Button(txt_display_updated_frame, text="Delete selected txt file", command=lambda:delete_selected_listbox_items(txt_updated_path_name_listbox), state=tk.DISABLED)
# -------------------------------------------------------------------------------------------#
# Adding widgets into the window

# Frame packing
xpad = 10
ypad = 20

txt_input_frame.pack(side='top', fill="both", expand=False, padx=xpad, pady=10)
txt_display_frame.pack(side='top', fill="both", expand=True, padx=xpad, pady=10)
txt_path_files_text_frame.pack(side='left', fill="both", expand=False, padx=xpad, pady=ypad)
txt_path_files_entry_frame.pack(side='left', fill="both", expand=True, padx=xpad, pady=ypad)
txt_path_files_entry_button_frame.pack(side='left', fill="both", expand=False, padx=xpad, pady=ypad)
txt_display_existing_frame.pack(side='left', fill="both", expand=True, padx=xpad, pady=ypad)
txt_display_updated_frame.pack(side='left', fill="both", expand=True, padx=xpad, pady=ypad) 

# first frame widgets
txt_root_directory_label.pack(side='top', fill="both", expand=True, padx=xpad)
txt_result_file_names_label.pack(side='top', fill="both", expand=True, padx=xpad)
txt_exisitng_name_label.pack(side='top', fill="both", expand=True, padx=xpad)
txt_replacing_name_label.pack(side='top', fill="both", expand=True, padx=xpad)
txt_final_folder_path_label.pack(side='top', fill="both", expand=True, padx=xpad)


txt_root_directory_entry.pack(side='top', fill="both", expand=True, padx=xpad, pady=ypad)
txt_result_file_names_entry.pack(side='top', fill="both", expand=True, padx=xpad, pady=ypad)
txt_exisitng_name_entry.pack(side='top', fill="both", expand=True, padx=xpad, pady=ypad)
txt_replacing_name_entry.pack(side='top', fill="both", expand=True, padx=xpad, pady=ypad)
txt_final_folder_path_entry.pack(side='top', fill="both", expand=True, padx=xpad, pady=ypad)

txt_root_dir_browse_button.pack(side='top', expand=False, padx=xpad, pady=ypad)
txt_dest_dir_browse_button.pack(side='bottom', expand=False, padx=xpad, pady=ypad)

#2nd frame widgets
txt_existing_file_label.pack(side='top', fill="both", expand=False, padx=xpad)
txt_existing_file_label.config(anchor="center")
txt_updated_file_label.pack(side='top', fill="both", expand=False, padx=xpad)
txt_updated_file_label.config(anchor="center")
txt_show_existing_button.pack(side='top', fill="both", expand=False, padx=xpad, pady=10)
txt_show_updated_button.pack(side='top', fill="both", expand=False, padx=xpad, pady=10) 
txt_exisitng_path_name_listbox.pack(side='top', fill="both", expand=True, padx=xpad, pady=(5,1))
txt_updated_path_name_listbox.pack(side='top', fill="both", expand=True, padx=xpad, pady=5)
txt_delete_new_files_button.pack(side='top', fill="both", expand=False, padx=xpad) 
# ---------- tabs 2 Widgets  ---------- #

# 2. Creates subframes inside fame above to allow grouping of other widgets
input_frame = tk.LabelFrame(frame, text="Inputs")
path_files_frame = tk.LabelFrame(input_frame, text="[STEP 1.] File Inputs", width=300, height=400)
path_files_input_frame = tk.Frame(path_files_frame)
path_files_text_frame = tk.Frame(path_files_input_frame)
path_files_entry_frame = tk.Frame(path_files_input_frame)
path_files_browse_frame = tk.Frame(path_files_input_frame)
path_files_scenario_frame = tk.LabelFrame(path_files_frame, text="[STEP 2.] Scenario Parameters (Please select keywords from txt file name)")
excel_tab_list_frame = tk.LabelFrame(input_frame, text="[STEP 3.] List of Excel Tabs (Worksheets)")
output_frame = tk.LabelFrame(frame, text="[STEP 4. Output")
scenario_years_frame = tk.Frame(path_files_scenario_frame)
scenario_state_frame = tk.Frame(path_files_scenario_frame)
scenario_period_frame = tk.Frame(path_files_scenario_frame)

# 3. Creates labels widgets
error_msg = tk.StringVar(value="Click Show Worksheets to verify valid path") # uses string var which gets updated basd on inputs
reminder_msg = tk.StringVar(value="Please select all parameters and worksheets\n before pressing the 'Select Worksheets' button below")
excel_path_label = ttk.Label(path_files_text_frame, text="Full Excel Path")
root_directory_label = ttk.Label(path_files_text_frame, text="Full Folder Path")
result_file_names_label = ttk.Label(path_files_text_frame, text="Raw Results File name")
delimiter_label = ttk.Label(path_files_text_frame, text="Delimiter")
excel_path_status_label = ttk.Label(excel_tab_list_frame, textvariable=error_msg)
user_reminder_label = ttk.Label(excel_tab_list_frame, textvariable=reminder_msg)

scenario_years_label = tk.Label(scenario_years_frame, text="Select Scenario Years")
scenario_state_label = tk.Label(scenario_state_frame, text="Select Scenario State")
scenario_period_label = tk.Label(scenario_period_frame, text="Select Scenario Period")

# 4. Creates pre-defined messages in the user input entries to give directions
excel_string = tk.StringVar(value='Enter full excel path file with no "" marks')
# root_string = tk.StringVar(value='Please enter full folder path file cotaining all results with no "" marks')
# raw_results_string = tk.StringVar(value="Enter a common results file name. eg.Node Results.att")
delimiter_string = tk.StringVar(value="Enter a delimiter to seperate the results")

# 5. Creates a listbox widget which will display exisitng excel tabs from user input
excel_worksheet_listbox = tk.Listbox(excel_tab_list_frame, selectmode="extended")
scenario_period_input_listbox = tk.Listbox(scenario_period_frame, selectmode="multiple", height=5)

# 6. Creates Entry widgets which allows users to type something
entry_width = 50
excel_path_entry = ttk.Entry(path_files_entry_frame, textvariable=excel_string, width=entry_width)
root_directory_entry = ttk.Entry(path_files_entry_frame, textvariable=txt_final_folder_path_string, width=entry_width)
result_file_names_entry = ttk.Entry(path_files_entry_frame, textvariable=txt_result_file_names_string, width=entry_width)
delimiter_entry = ttk.Entry(path_files_entry_frame, textvariable=delimiter_string, width=entry_width)

# 7. Creates a text box for user to add 
scenario_years_input_text = tk.Text(scenario_years_frame, height=5, width=15)
scenario_state_input_text = tk.Text(scenario_state_frame, height=5, width=15)

# 8. Creates button widgets which calls functions when pressed
    # lambda function allows you to input widgets as arguments 
root_dir_browse_button = tk.Button(path_files_browse_frame, text="Browse Excel", width=15, command=lambda:browse(excel_string, False))
dest_dir_browse_button = tk.Button(path_files_browse_frame, text="Browse Destination", width=15, command=lambda:browse(txt_final_folder_path_string))

show_valid_worksheets_button = ttk.Button(excel_tab_list_frame, text="Show Worksheets", command=lambda:get_worksheet_names(excel_string,excel_worksheet_listbox))
select_worksheet_button = ttk.Button(excel_tab_list_frame, text="Select Workheets", command=lambda:get_selected_worksheet_no(excel_worksheet_listbox, select_worksheet_button), state=tk.DISABLED)
update_excel_button = ttk.Button(output_frame, text="Update Excel", command=lambda:update_func(excel_worksheet_listbox), state=tk.DISABLED)

scenario_years_button = ttk.Button(scenario_years_frame, text="Select years", command=lambda:convert_string_to_list(scenario_years, scenario_years_input_text, scenario_years_button, "scenario_years_button"))
scenario_state_button = ttk.Button(scenario_state_frame, text="Select state", command=lambda:convert_string_to_list(scenario_state, scenario_state_input_text, scenario_state_button, "scenario_state_button"))
scenario_period_button = ttk.Button(scenario_period_frame, text="Select periods", command=lambda:assign_scenario_period_to_list(scenario_period_input_listbox, scenario_period_button))

#---------------------------------------------------------#
# ---------- Placing all the widgets in window ---------- #
#---------------------------------------------------------#

# 1. place the frames into the window
xpad = 10
ypad = 20
    # main two frames
input_frame.pack(side='top', fill="both", expand=True, padx=xpad, pady=10)
output_frame.pack(side='top', fill="both", expand=False, padx=xpad, pady=10)
    # two sub frames inside the input frame
path_files_frame.pack(side='left', fill="both", expand=True, padx=xpad, pady=10)
path_files_frame.pack_propagate(False) # stops the frames from scaling
excel_tab_list_frame.pack(side='left', fill="both", expand=True, padx=xpad, pady=10)
excel_tab_list_frame.pack_propagate(False) 
    # sub-sub frames inside the "File Inputs" frame
path_files_input_frame.pack(side='top', fill="both", expand=False)
path_files_scenario_frame.pack(side='top', fill="both", expand=True)
path_files_text_frame.pack(side='left', fill="both", expand=False, padx=xpad, pady=ypad)
path_files_entry_frame.pack(side='left', fill="both", expand=True, padx=xpad, pady=ypad)
path_files_browse_frame.pack(side='left', fill="both", expand=False, padx=xpad, pady=ypad)
    # sub-sub frames inside the "scenario paramters" frame
scenario_years_frame.pack(side='left', fill="both", expand=True, padx=xpad)
scenario_state_frame.pack(side='left', fill="both", expand=True, padx=xpad)
scenario_period_frame.pack(side='left', fill="both", expand=True, padx=xpad)

# 2. place the labels into the label frame
excel_path_label.pack(side='top', fill="both", pady=ypad)
root_directory_label.pack(side='top', fill="both", pady=ypad)
result_file_names_label.pack(side='top', fill="both", pady=ypad)
delimiter_label.pack(side='top', fill="both", pady=ypad)

# 3. place the entries into the entry frame
excel_path_entry.pack(side='top', fill="both", pady=ypad)
root_directory_entry.pack(side='top', fill="both", pady=ypad)
result_file_names_entry.pack(side='top', fill="both", pady=ypad)
delimiter_entry.pack(side='top', fill="both", pady=ypad)


root_dir_browse_button.pack(side='top', expand=False, pady=18)
dest_dir_browse_button.pack(side='top', expand=False, pady=18)

# 4. place the widgets into the list frame
excel_path_status_label.pack(side='top', fill="both", pady=9)
excel_path_status_label.config(anchor="center")
show_valid_worksheets_button.pack(side='top', fill="both", padx=xpad)
excel_worksheet_listbox.pack(side='top', fill="both", expand=True, padx=xpad, pady=5)
user_reminder_label.pack(side='top', fill="both", pady=0)
user_reminder_label.config(anchor="center")
select_worksheet_button.pack(side='top', fill="both", padx=xpad, pady=10)

# 5. place the widgets into the scenario param frame
scenario_years_label.pack(side='top', pady=(0, ypad))
scenario_state_label.pack(side='top', pady=(0, ypad))
scenario_period_label.pack(side='top', pady=(0, ypad), anchor=tk.N)

scenario_years_input_text.pack(side='top', fill="both", expand=True)
scenario_state_input_text.pack(side='top', fill="both", expand=True)

scenario_period_input_listbox.pack(side='top', fill="both", expand=True)
scenario_period_input_listbox.insert(1, "AM")
scenario_period_input_listbox.insert(1, "PM")

scenario_years_button.pack(side='top', fill="both", pady=(ypad + 5, 5))
scenario_state_button.pack(side='top', fill="both", pady=(ypad + 5, 5))
scenario_period_button.pack(side='top', fill="both", pady=(ypad + 5, 5))

# 6. place the widgets into the update frame
update_excel_button.pack(side='top', fill="both", expand=True, padx=xpad, pady=ypad)

#---------------------------------------#
# ---------- Runs the window ---------- #
#---------------------------------------#

window.mainloop()