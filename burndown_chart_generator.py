import collections
import datetime
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import os
from tkinter import filedialog
import sys
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, messagebox
from tkcalendar import Calendar
import logging
import warnings

warnings.filterwarnings("ignore", message="Workbook contains no default style, apply openpyxl's default")


logging.basicConfig(filename='script.log', level=logging.DEBUG)
logging.debug('This is a debug message')
logging.info('This is an info message')
logging.warning('This is a warning message')
logging.error('This is an error message')
logging.critical('This is a critical message')

script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)
# ==========================================
#  Title:  Burndown Chart Generator
#  Author: Yifei Wang
#  yifei.wang@valeo.com
#  Date:   22/03/2023
# ==========================================
#################################################################################
################################### README ######################################
# this script is aim to collect data from an exported Excel sheet and process   #
# the data to generate the burndown chart:                                      #
# x-axix: selected date range (from start date to end date)                     #
# y-axis: ideal and actuol remaining tasks                                      #
# bar: completed task within the select date range                              #
#################################################################################
############################# Input for Script ##################################
# The input for the script is an excel sheet which needs to be exported from    #
# TeamForge manually. Please ensure the following items are included:           #
#                          1. Artifact ID                                       #
#                          2. Due Date                                          #
#                          3. Last Status Change                                #
#                          4. Status                                            #
#                          5. Planned For                                       #
#################################################################################
# You need to save this excel sheet locally and pass the path to the pop-up     #
# window.                                                                       #
############################# Output for Script ##################################
# The script generate a burndown chart as a pdf format, you can save the figure #
# to your local drive by clicking the button SAVE                               #
#################################################################################

def get_date(text):
    def center_window(window, width, height):
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()

        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)

        window.geometry(f'{width}x{height}+{x}+{y}')

    def cal_done():
        top.withdraw()
        root.quit()

    root = tk.Tk()
    root.withdraw()  # keep the root window from appearing

    top = tk.Toplevel(root)

    # Set the desired width and height for the Toplevel window
    width = 1200
    height = 900
    center_window(top, width, height)

    cal = Calendar(top,
                   font="Arial 16", selectmode='day',
                   cursor="hand1")
    cal.pack(fill="both", expand=True)
    tk.Button(top, text=text, command=cal_done, font=("Arial", 16), background='green', foreground='black',
              activebackground='green', activeforeground='white', bd=2, relief='raised').pack()

    selected_date = None
    root.mainloop()

    return cal.selection_get()
def convertDate2Datetime(date):

    if isinstance(date, str):
        closedDateYear = int(date.split('/')[2][:4])
        closedDateMonth = int(date.split('/')[0])
        closedDateDay = int(date.split('/')[1])
        date = datetime.date(closedDateYear, closedDateMonth, closedDateDay)
    else:
        date = ''
    return date

def generate_data_list(start_date,end_date):

    return pd.date_range(start_date,end_date, freq='d')

def read_excel_sheet():

    file_path = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select the path of exported Excel from TeamForge",
                                           filetypes=[("Excel files", "*.xlsx")])
    data = pd.read_excel(file_path)

    # Drop rows that don't start with 'PI' and don't contain 'Sprint'
    excel_df = data[data['Planned For'].str.startswith('PI') & data['Planned For'].str.contains('Sprint')]

    return excel_df

def filter_data_by_regex(data, pi_val, team_org_val, sprint_val):
    regex_pattern = f'.*{pi_val}.*{team_org_val}.*Sprint {sprint_val}.*'
    filtered_data = data[data['Planned For'].str.match(regex_pattern)]
    if filtered_data.empty == True:
        messagebox.showwarning("Warning", "Can not find your input in column 'Planned For'")
        sys.exit()
    return filtered_data

# Create a simple GUI using Tkinter to get user input
def submit():
    global res1, res2, pi_val, team_org_val, sprint_val, priority_val
    pi_val = pi_input.get()
    team_org_val = team_org_input.get()
    sprint_val = sprint_input.get()
    priority_val = priority_input.get()
    res1 = filter_data_by_regex(excel_df,pi_val,team_org_val,sprint_val)
    res2 = res1[res1['Priority'] == int(priority_val)]
    submit_flag.set('1')
    root.quit()

def create_input_gui():
    global pi_input, team_org_input, sprint_input, priority_input, root, submit_flag
    root = tk.Tk()
    root.title("User Inputs for generating Burndown-Chart:")

    # Set the window size
    window_width = 600
    window_height = 150

    # Calculate the center position
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x_position = int((screen_width / 2) - (window_width / 2))
    y_position = int((screen_height / 2) - (window_height / 2))

    # Set the window geometry
    root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

    # Set font size for better readability
    font = tkfont.Font(size=17)

    submit_flag = tk.StringVar()
    submit_flag.set('0')

    tk.Label(root, text="PI").grid(row=0, column=0)
    pi_input = tk.Entry(root, font=font, width=50, justify='center')
    pi_input.grid(row=0, column=1)

    tk.Label(root, text="Team Organization").grid(row=1, column=0)
    team_org_input = tk.Entry(root, font=font, width=50, justify='center')
    team_org_input.grid(row=1, column=1)

    tk.Label(root, text="Sprint").grid(row=2, column=0)
    sprint_input = tk.Entry(root, font=font, width=50, justify='center')
    sprint_input.grid(row=2, column=1)

    tk.Label(root, text="Priority").grid(row=3, column=0)
    priority_input = tk.Entry(root, font=font, width=50, justify='center')
    priority_input.grid(row=3, column=1)
    root.grid_rowconfigure(4, minsize=20) # Add an empty row with a height of 20 pixels
    button_width = 10  # Set the button width
    tk.Button(root, text="Submit", command=lambda: submit(), font=font, bg='green', width=button_width).grid(row=5,
                                                                                                             column=0,
                                                                                                             columnspan=2)
    root.grid_columnconfigure(0, weight=1)  # Add padding to the left of the button
    root.grid_columnconfigure(1, weight=1)  # Add padding to the right of the button
    root.mainloop()

def sheet_data_processing(excel_df):

    id_list = []
    due_date_list = []

    for k in excel_df.index:
        date = excel_df.loc[k, 'Due Date']
        if isinstance(date, str):  # drop all items without containing a due date
            if convertDate2Datetime(date) < start_date or convertDate2Datetime(date) > end_date:
                excel_df = excel_df.drop(k)
            else:
                id_list.append(excel_df.loc[k, 'Artifact ID'])
                due_date_list.append(excel_df.loc[k, 'Due Date'])
        else:
            excel_df = excel_df.drop(k)

    due_date_id_list = [(id, cw) for id, cw in zip(id_list, due_date_list)]

    tasks_per_due_date_dict = collections.defaultdict(list)

    for j, k in due_date_id_list:
        tasks_per_due_date_dict[k].append(j)

    # list all CWs
    due_date_id_list_keys = tasks_per_due_date_dict.keys()
    num_tasks_per_due_date_list = []  # a list count the artifacts for each date
    for i in tasks_per_due_date_dict.keys():
        num_tasks_per_due_date_list.append(len(tasks_per_due_date_dict[i]))

    df = pd.DataFrame(np.transpose(np.array([list(due_date_id_list_keys), num_tasks_per_due_date_list])),
                      columns=['due date', 'planned'])
    df = df.sort_values(by=['due date'])  # sort the values by CWs

    # create the other 3 columns
    df['done as plan'] = 0
    df['ideal remaining tasks'] = 0
    df['actual remaining tasks'] = 0

    for i in excel_df.index:
        if excel_df.loc[i, 'Status'] == 'Closed' and convertDate2Datetime(
                excel_df.loc[i, 'Last Status Change']) <= convertDate2Datetime(excel_df.loc[i, 'Due Date']):
            excel_df.loc[
                i, 'completed as planned'] = True  # column 'completed as planned' shows Boolean which displays if the tasks completed
            due_date_nr = excel_df.loc[i, 'Due Date']
            df.loc[df[
                       'due date'] == due_date_nr, 'done as plan'] += 1  # column 'done as plan' shows the count of tasks completed as plan
        else:
            excel_df.loc[i, 'completed as planned'] = False

    df2 = df.reset_index(drop=True)  # Update index after sorting data-frame

    global date_range
    date_range = generate_data_list(start_date, end_date)

    data3 = np.zeros(len(date_range))

    df3 = pd.DataFrame([date_range, data3, data3, data3, data3])
    df3 = df3.transpose()

    df3.columns = ['date', 'planned', 'done as plan', 'ideal remaining tasks', 'actual remaining tasks']

    for i in df2.index:
        df2.loc[df2.index[i], 'ideal remaining tasks'] = sum(df2['planned'].astype(int)) - sum(
            df2['planned'][0:i + 1].astype(int))

    # Attempt to convert 'due date' column to datetime64, handle exceptions
    try:
        df2['due date'] = pd.to_datetime(df2['due date'], errors='raise')
        # print("Successfully converted 'due date' column to datetime64")
    except Exception as e:
        print(f"Error occurred while converting 'due date' column: {e}")

    global df_idx_list
    df_idx_list = []
    for i in df3['date'].index:
        for j in df2['due date'].index:
            if df3['date'].values[i] == df2['due date'].values[j]:
                df_idx_list.append(i)
                df3.loc[i, 'planned'] = df2['planned'].values[j]
                df3.loc[i, 'done as plan'] = df2['done as plan'].values[j]
                df3.loc[i, 'ideal remaining tasks'] = df2['ideal remaining tasks'].values[j]
    df3['done_at_this_day'] = 0

    for i in excel_df.index:
        if excel_df.loc[i, 'Status'] == 'Closed':
            for j in df3.index:
                if convertDate2Datetime(excel_df.loc[i, 'Last Status Change']) == df3['date'].values[j].astype(
                        'M8[D]').astype('O'):
                    df3.loc[j, 'done_at_this_day'] += 1

    for j in df3['date'].index:
        df3.loc[j, 'actual remaining tasks'] = sum(df3['planned'].astype(int)) - sum(df3['done_at_this_day'][:j + 1])

    df3.loc[0, 'ideal remaining tasks'] = sum(df3['planned'].astype(int))

    return df3

def plot(df):
    actual_plot_flag = True
    ideal_plot_flag = True
    now_date_compare_date_range = 0 # -1: eariler than the range; 0:in the range ; 1: later than the range
    fig, ax = plt.subplots()  # Create a figure containing a single axes
    now_date = datetime.date.today()
    datetime_obj = datetime.datetime.combine(now_date, datetime.datetime.min.time())
    datetime64_obj = np.datetime64(datetime_obj)
    if datetime64_obj in date_range:
        now_date_idx = date_range.get_loc(datetime64_obj)+1
    else:
        now_date_idx = len(date_range)
    df_idx_list.insert(0, 0)

    if df['actual remaining tasks'].isna().any():
        print("The 'actual remaining tasks' column contains NaN values. Skipping plot.")
        actual_plot_flag = False

    if df['ideal remaining tasks'].isna().any():
        print("The 'ideal remaining tasks' column contains NaN values. Skipping plot.")
        ideal_plot_flag = False
    else:
        ax.plot(date_range[df_idx_list], df['ideal remaining tasks'][df_idx_list], label='Ideal')

    if datetime64_obj not in date_range:
         if datetime64_obj < date_range[0]:
            now_date_compare_date_range = -1
         elif datetime64_obj > date_range[0]:
            now_date_compare_date_range = 1
    else:
        now_date_compare_date_range = 0

    if actual_plot_flag == True:
        if datetime64_obj in date_range:
            ax.plot(date_range[:now_date_idx], df['actual remaining tasks'][:now_date_idx], label='Actual')
        elif now_date_compare_date_range == 1:
            ax.plot(date_range, df['actual remaining tasks'], label='Actual')

    # add annotations to the actual curve
    if actual_plot_flag == True and now_date_compare_date_range != -1:
        for i, val in enumerate(df['actual remaining tasks'][:now_date_idx]):
            ax.annotate(str(val), xy=(date_range[i], val), ha='center', va='bottom', fontsize=7)

    # add annotations to the ideal curve
    if ideal_plot_flag == True:
        for i, val in enumerate(df['ideal remaining tasks'][df_idx_list]):
            ax.annotate(str(val), xy=(date_range[df_idx_list][i], val), ha='center', va='bottom', fontsize=7)

    plt.xticks(rotation=90)

    bars = ax.bar(date_range, df['done_at_this_day'], label='Completed Tasks')
    ax.bar_label(bars, fontsize=7)
    plt.legend(loc='center left', fontsize=7)
    plt.xticks(date_range, fontsize=7)
    plt.yticks(fontsize=7)
    if plot_prio_chart == False:
        title = "Burndown Chart" + ' ' + 'PI ' + pi_input.get() + ' ' + team_org_input.get() + ' Sprint ' +sprint_input.get()
    else:
        title = "Burndown Chart" + ' ' + 'PI ' + pi_input.get() + ' ' + team_org_input.get() + ' Sprint ' +sprint_input.get() + ' Prio: ' + priority_val
    filename = title.replace('>', '-').replace(' ', '_') + '.pdf'
    plt.xlabel('Days')
    plt.ylabel('Remaining Tasks')
    plt.title(title)
    plt.grid()
    plt.savefig(filename, format="pdf", bbox_inches="tight")
    plt.show()

def main():

    global start_date, end_date, plot_prio_chart
    plot_prio_chart = False
    start_date = get_date('Select Start Date')
    end_date = get_date('Select End Date')
    global excel_df
    excel_df= read_excel_sheet()
    create_input_gui()
    submit()
    df1 = sheet_data_processing(res1)
    plot(df1)
    plot_prio_chart = True
    df2 = sheet_data_processing(res2)
    plot(df2)

if __name__ == '__main__':
    main()
