import collections
import datetime
import re
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import os
from tkinter import filedialog
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QLineEdit
from PyQt5.QtGui import QFont
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import Calendar, DateEntry
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
class MyWindow(QWidget):
    def __init__(self):

        super().__init__()

        self.initUI()

    def initUI(self):

        vbox = QVBoxLayout()

        self.label = QLabel('Enter the full sprint structure')
        font = QFont('Arial', 15, QFont.Bold)
        self.label.setFont(font)
        vbox.addWidget(self.label)

        self.info_label = QLabel('e.g.PI 1 > Lidar OS > Integration Team > Sprint 3')
        self.info_label.setFont(font)
        vbox.addWidget(self.info_label)

        vbox.setSpacing(10)

        self.textbox = QLineEdit(self)
        self.textbox.setFixedHeight(50)
        vbox.addWidget(self.textbox)
        self.textbox.setFont(QFont('Arial', 15))

        self.button = QPushButton('Submit')
        self.button.setFont(font)
        self.button.clicked.connect(self.return_text)
        vbox.addWidget(self.button)

        self.setLayout(vbox)
        self.resize(500, 200)
        self.show()

    def return_text(self):

        text: str = self.textbox.text()
        print('Entered text:', text)
        global enter_text
        enter_text = text
        self.close()
def get_date(text):

    def cal_done():
        top.withdraw()
        root.quit()

    root = tk.Tk()
    root.withdraw() # keep the root window from appearing

    top = tk.Toplevel(root)

    cal = Calendar(top,
                   font="Arial 14", selectmode='day',
                   cursor="hand1")
    cal.pack(fill="both", expand=True)
    ttk.Button(top, text=text, command=cal_done).pack()

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
        # ax.plot(date_range[df_idx_list], df['ideal remaining tasks'][df_idx_list], label='Ideal')

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
    title = "Burndown Chart" + ' ' + enter_text
    filename = title.replace('>', '-').replace(' ', '_') + '.pdf'
    plt.xlabel('Days')
    plt.ylabel('Remaining Tasks')
    plt.title(title)
    plt.grid()
    plt.savefig(filename, format="pdf", bbox_inches="tight")
    plt.show()
    print('Success!')

def read_excel_sheet():

    file_path = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select the path of exported Excel from TeamForge",
                                           filetypes=[("Excel files", "*.xlsx")])

    return pd.read_excel(file_path)


def enter_sprint_struct():
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec_()

def valid_input(input,excel_df):
    if enter_text not in list(excel_df['Planned For']):
        root = tk.Tk()
        root.withdraw()
        messagebox.showwarning("Warning", "Can not find your input in column 'Planned For'")
        root.destroy()
        sys.exit()


def sheet_data_processing(excel_df):
    regex_pattern = re.compile(enter_text)

    idx_list = []
    # items used for generating burndown chart from the exported excel sheet from TeamForge
    cols = ['Due Date', 'Last Status Change', 'Status', 'Planned For']

    for k in np.arange(0, excel_df.shape[0]):
        if isinstance(excel_df.loc[k, 'Planned For'], str):  # drop all empty cells
            if not re.search(regex_pattern, excel_df.loc[k, 'Planned For']):  # drop all items not match regex_pattern
                excel_df = excel_df.drop(k)
                idx_list.append(k)
        else:
            excel_df = excel_df.drop(k)
            idx_list.append(k)

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

    # df2['due date'] = df2['due date'].astype('datetime64')
    # Attempt to convert 'due date' column to datetime64, handle exceptions
    try:
        df2['due date'] = pd.to_datetime(df2['due date'], errors='raise')
        print("Successfully converted 'due date' column to datetime64")
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
def main():

    global start_date, end_date
    start_date = get_date('Select Start Date')
    end_date = get_date('Select End Date')

    excel_df= read_excel_sheet()

    enter_sprint_struct()

    valid_input(enter_text,excel_df)

    df3 = sheet_data_processing(excel_df)

    plot(df3)


if __name__ == '__main__':
    main()
