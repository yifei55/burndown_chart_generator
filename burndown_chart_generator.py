import collections
import datetime
import re
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import os
from tkinter import filedialog
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QComboBox, QLabel, QPushButton, QLineEdit
from PyQt5.QtGui import QFont
from PyQt5.QtCore import pyqtSignal
# ==========================================
#  Title:  Burndown Chart Generator
#  Author: Yifei Wang
#  Date:   03/03/2023
# ==========================================
#################################################################################
################################### README ######################################
# this script is aim to collect data from an exported Excel sheet and process   #
# the data to generate the burndown chart (x-axix: calendar week, y-axis: ideal #
# and actual remaining tasks, bar: completed task within this calendar week)    #
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
# You need to save this excel sheet locally and pass the path to the variable   #
# 'excel_df'.                                                                   #
############################# Output for Script ##################################
# The script generate a burndown chart as a pdf format, you can save the figure #
# to your local drive by clicking the button SAVE                               #
#################################################################################

######################### Project-Depended Variables ############################

# class CascadingDropdowns(QWidget):
#
#     submit_clicked = pyqtSignal(str, str, str)
#     # Initialize the selected values as None
#
#     def __init__(self):
#         super().__init__()
#
#         self.init_ui()
#         self.level1_selected = None
#         self.level2_selected = None
#         self.level3_selected = None
#     def init_ui(self):
#         self.setWindowTitle('Cascading Dropdowns')
#         self.resize(400, 200)
#
#         layout = QVBoxLayout()
#
#         # Create level 1 dropdown
#         self.level1_dropdown = QComboBox()
#         self.level1_dropdown.addItems(['China', 'USA', 'Germany'])
#         self.level1_dropdown.currentIndexChanged.connect(self.update_level2_dropdown)
#
#         # Create level 2 dropdown
#         self.level2_dropdown = QComboBox()
#
#         # Create level 3 standalone dropdown
#         self.level3_dropdown = QComboBox()
#         self.level3_dropdown.addItems(['a', 'b', 'c'])
#
#         # Create submit button
#         self.submit_button = QPushButton('Submit')
#         self.submit_button.clicked.connect(self.pass_selected_values)
#
#         # Add widgets to layout
#         layout.addWidget(QLabel('Level 1:'))
#         layout.addWidget(self.level1_dropdown)
#         layout.addWidget(QLabel('Level 2:'))
#         layout.addWidget(self.level2_dropdown)
#         layout.addWidget(QLabel('Level 3:'))
#         layout.addWidget(self.level3_dropdown)
#         layout.addWidget(self.submit_button)
#
#         self.setLayout(layout)
#         self.update_level2_dropdown()
#
#         # Apply styles
#         self.apply_styles()
#
#     def apply_styles(self):
#         style = """
#         QWidget {
#             font-family: "Arial";
#             font-size: 14px;
#             background-color: #f0f0f0;
#         }
#
#         QLabel {
#             font-weight: bold;
#             color: #303030;
#         }
#
#         QComboBox {
#             background-color: #ffffff;
#             border: 1px solid #303030;
#             padding: 2px;
#         }
#
#         QComboBox::drop-down {
#             border: 0px;
#         }
#
#         QComboBox::down-arrow {
#             image: url('path/to/your/arrow/image.png');
#             width: 14px;
#             height: 14px;
#         }
#
#         QPushButton {
#             background-color: #303030;
#             color: #ffffff;
#             padding: 5px;
#             border-radius: 3px;
#         }
#
#         QPushButton:hover {
#             background-color: #505050;
#         }
#         """
#
#         self.setStyleSheet(style)
#
#     def update_level2_dropdown(self):
#         level1_item = self.level1_dropdown.currentText()
#         level2_items = []
#
#         if level1_item == 'China':
#             level2_items = ['Shiyan', 'Wuhan']
#         elif level1_item == 'USA':
#             level2_items = ['Boston', 'LA']
#         elif level1_item == 'Germany':
#             level2_items = ['Koln', 'Stuttgart']
#
#         self.level2_dropdown.clear()
#         self.level2_dropdown.addItems(level2_items)
#
#     def pass_selected_values(self):
#         level1_value = self.level1_dropdown.currentText()
#         level2_value = self.level2_dropdown.currentText()
#         level3_value = self.level3_dropdown.currentText()
#
#         print(f"Level 1: {level1_value}, Level 2: {level2_value}, Level 3: {level3_value}")
#         self.submit_clicked.emit(level1_value, level2_value, level3_value)
#         self.close()

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
        self.resize(400, 200)
        self.show()

    def return_text(self):
        text: str = self.textbox.text()
        print('Entered text:', text)
        global enter_text
        enter_text = text

        # Here you can call your function and pass the entered text as an argument.
        self.close()


def get_date(text):
    import tkinter as tk
    from tkinter import ttk
    from tkcalendar import Calendar, DateEntry

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

#################################################################################

def convertDate2Datetime(date):
    '''
    convert the date to calendar week
    :param date: the raw date parsed from excel
    :return: the calendar week after conversion
    '''
    if isinstance(date, str):
        closedDateYear = int(date.split('/')[2][:4])
        closedDateMonth = int(date.split('/')[0])
        closedDateDay = int(date.split('/')[1])
        date2 = datetime.date(closedDateYear, closedDateMonth, closedDateDay)
    else:
        date2 = ''
    return date2

def generate_data_list(start_date,end_date):

    return pd.date_range(start_date,end_date, freq='d')

def on_submit_clicked(level1, level2, level3):
    print(f"Selected values: Level 1: {level1}, Level 2: {level2}, Level 3: {level3}")
    global a1, b1, c1
    a1 = level1
    b1 = level2
    c1 = level3

def main():
    start_date = get_date('Select Start Date')
    end_date = get_date('Select End Date')

    file_path = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select the path of exported Excel from TeamForge",
                                           filetypes=[("Excel files", "*.xlsx")])

    # read the excel from local and convert it to DataFrame
    excel_df = pd.read_excel(file_path)

    # create a pandas series
    s = pd.Series(excel_df['Planned For'].unique())
    s_filtered_1 = [x for x in s if x is not None and not pd.isna(x)]
    s_filtered_2 = [s_filtered_1 for s_filtered_1 in s_filtered_1 if 'PI 1' in s_filtered_1]

    s_filtered_3 = [x.replace('PI 1 > ', '') if 'PI 1 > ' in x else x for x in s_filtered_2]

    split_series = s.str.split('>', expand=True)
    list_1 = list(set(split_series[0].tolist()))
    list_2 = list(set(split_series[1].tolist()))
    list_3 = list(set(split_series[2].tolist()))


    filtered_list_1 = [x for x in list_1 if x is not None and not pd.isna(x)]
    filtered_list_2 = [x for x in list_2 if x is not None and not pd.isna(x)]
    filtered_list_3 = [x for x in list_3 if x is not None and not pd.isna(x)]


    modified_list_1 = [x.strip() for x in filtered_list_1]
    modified_list_2 = [x.strip() for x in filtered_list_2]
    modified_list_3 = [x.strip() for x in filtered_list_3]


    app = QApplication(sys.argv)
    # window = CascadingDropdowns()
    # window.submit_clicked.connect(on_submit_clicked)
    window = MyWindow()
    window.show()
    app.exec_()

    planned_for_group_lv1 = 'Product'
    planned_for_group_lv2 = 'HIL'
    sprint_nr = 'Sprint 3'
    # use regular expression to filter out the keywords as Lidar OS and Sprint 3, this can be adjusted based on needs
    # regex_pattern = pattern = r'^(?=.*planned_for_group_lv1\s*=\s*{0})(?=.*planned_for_group_lv2\s*=\s*{1})(?=.*sprint_nr\s*=\s*{2}).*$'.format(planned_for_group_lv1, planned_for_group_lv2, sprint_nr)
    # regex_pattern = re.compile(r'^(?=.*?\bProduct\b)(?=.*?\bHIL\b)(?=.*?\bSprint 3\b).*$')
    # regex_pattern = re.compile(r'^(?=.*?\b + planned_for_group_lv1 + \b)(?=.*?\b + planned_for_group_lv2 + \b)(?=.*?\b + sprint_nr + \b).*$')

    regex_pattern = re.compile(enter_text)


    # filter out all LiDAR OS Sprint 3 tasks
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
    # CW_list = []
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

    now_date = datetime.date.today()
    # get the current CW (week_num) because the CW of actual remaining tasks shall not newer than the current CW
    year, week_num, day_of_week = now_date.isocalendar()

    df2 = df.reset_index(drop=True)  # Update index after sorting data-frame

    date_range = generate_data_list(start_date,end_date)

    data3 = np.zeros(len(date_range))

    df3 = pd.DataFrame([date_range, data3, data3, data3, data3])
    df3 = df3.transpose()

    df3.columns=['date', 'planned', 'done as plan', 'ideal remaining tasks', 'actual remaining tasks']

    for i in df2.index:
        df2.loc[df2.index[i], 'ideal remaining tasks'] = sum(df2['planned'].astype(int)) - sum(
            df2['planned'][0:i + 1].astype(int))

    df2['due date'] = df2['due date'].astype('datetime64')

    df_idx_list = []
    for i in df3['date'].index:
        for j in df2['due date'].index:
            if df3['date'].values[i] == df2['due date'].values[j]:
                df_idx_list.append(i)
                df3.loc[i,'planned'] = df2['planned'].values[j]
                df3.loc[i,'done as plan'] = df2['done as plan'].values[j]
                df3.loc[i,'ideal remaining tasks'] = df2['ideal remaining tasks'].values[j]
    df3['done_at_this_day'] = 0

    for i in excel_df.index:
        if excel_df.loc[i, 'Status'] == 'Closed':
            for j in df3.index:
                if convertDate2Datetime(excel_df.loc[i, 'Last Status Change']) == df3['date'].values[j].astype('M8[D]').astype('O'):
                    df3.loc[j,'done_at_this_day']+=1

    for j in df3['date'].index:
        df3.loc[j, 'actual remaining tasks'] = sum(df3['planned'].astype(int)) - sum(df3['done_at_this_day'][:j+1])


    df3.loc[0,'ideal remaining tasks'] = sum(df3['planned'].astype(int))

    fig, ax = plt.subplots()  # Create a figure containing a single axes


    df_idx_list.insert(0,0)
    ax.plot(date_range, df3['actual remaining tasks'], label='Actual')
    ax.plot(date_range[df_idx_list], df3['ideal remaining tasks'][df_idx_list], label='Ideal')

    # add annotations to the actual curve
    for i, val in enumerate(df3['actual remaining tasks']):
        ax.annotate(str(val), xy=(date_range[i], val), ha='center', va='bottom', fontsize=7)

    # add annotations to the ideal curve
    for i, val in enumerate(df3['ideal remaining tasks'][df_idx_list]):
        ax.annotate(str(val), xy=(date_range[df_idx_list][i], val), ha='center', va='bottom', fontsize=7)

    plt.xticks(rotation=90)

    bars = ax.bar(date_range, df3['done_at_this_day'], label='Completed Tasks')
    ax.bar_label(bars, fontsize=7)
    plt.legend(loc='center left', fontsize=7)
    plt.xticks(date_range, fontsize=7)
    plt.yticks(fontsize=7)
    # title = "Burndown Chart" + ' ' + planned_for_group_lv1 + ' ' + planned_for_group_lv2 + ' ' + sprint_nr
    title = "Burndown Chart" + ' ' + enter_text
    filename = title.replace('>', '-').replace(' ', '_') + '.pdf'
    plt.xlabel('Days')
    plt.ylabel('Remaining Tasks')
    plt.title(title)
    plt.grid()
    plt.savefig(filename, format="pdf", bbox_inches="tight")
    plt.show()
    print('Success!')

if __name__ == '__main__':
    main()
    # app = QApplication(sys.argv)
    # window = CascadingDropdowns()
    # window.show()
    # sys.exit(app.exec_())


