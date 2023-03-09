import collections
import datetime
import re
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

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
# the following variables are adjustable based on your needs                    #

# read the excel from local and convert it to DataFrame
excel_df = pd.read_excel(
    r"C:\Python_script-20230222T094046Z-001\Python_script\proj5189-tracker30629-20230302-1613.xlsx")
planned_for_group = 'Lidar OS'
sprint_nr = 'Sprint 3'
# use regular expression to filter out the keywords as Lidar OS and Sprint 3, this can be adjusted based on needs
regex_pattern = r"(?=.*\bLidar\sOS\b)(?=.*\bSprint\s3\b).+"
#################################################################################

def convertDate2CWs(date):
    '''
    convert the date to calendar week
    :param date: the raw date parsed from excel
    :return: the calendar week after conversion
    '''
    if isinstance(date, str):
        closedDateYear = int(date.split('/')[2][:4])
        closedDateMonth = int(date.split('/')[0])
        closedDateDay = int(date.split('/')[1])
        date_calendarWeek = datetime.date(closedDateYear, closedDateMonth, closedDateDay).isocalendar()[1]
    else:
        date_calendarWeek = ''
    return date_calendarWeek

# filter out all LiDAR OS Sprint 3 tasks
idx_list = []
# items used for generating burndown chart from the exported excel sheet from TeamForge
cols = ['Due Date', 'Last Status Change', 'Status', 'Planned For']

for k in np.arange(0, excel_df.shape[0]):
    if isinstance(excel_df.loc[k, 'Planned For'], str): # drop all empty cells
        if not re.search(regex_pattern, excel_df.loc[k, 'Planned For']): # drop all items not match regex_pattern
            excel_df = excel_df.drop(k)
            idx_list.append(k)
    else:
        excel_df = excel_df.drop(k)
        idx_list.append(k)

# insert two new columns which need for further processing
for k in excel_df.index:
    excel_df.loc[k, 'due date CWs'] = convertDate2CWs(excel_df.loc[k, 'Due Date'])
    excel_df.loc[k, 'Last Status Change CWs'] = convertDate2CWs(excel_df.loc[k, 'Last Status Change'])

id_list = []
CW_list = []

for k in excel_df.index:
    date = excel_df.loc[k, 'Due Date']

    if isinstance(date, str):   # drop all items without containing a due date
        if int(date.split('/')[2][:4]) != 2023:   # drop all tasks not from 2023
            excel_df = excel_df.drop(k)
        else:
            id_list.append(excel_df.loc[k, 'Artifact ID'])
            CW_list.append(excel_df.loc[k, 'due date CWs'])
    else:
        excel_df = excel_df.drop(k)

CW_id_list = zip(id_list, CW_list)
# Create a list dynamically and store CWs matching with IDs
tasks_per_CWs_dict = collections.defaultdict(list)

# a CW list that each CW display all artifacts whose due date belong to:
# CW1: artf0001, artf0002
# CW2: artf0003, artf0004, artf0004...
# CW...
for j, k in CW_id_list:
    tasks_per_CWs_dict[k].append(j)

# list all CWs
cw_id_list_keys = tasks_per_CWs_dict.keys()
num_tasks_per_cw_list = [] # a list count the artifacts for each CWs
for i in tasks_per_CWs_dict.keys():
    num_tasks_per_cw_list.append(len(tasks_per_CWs_dict[i]))

df = pd.DataFrame(np.transpose(np.array([list(cw_id_list_keys), num_tasks_per_cw_list])), columns=['CWs', 'planned'])
df = df.sort_values(by=['CWs']) # sort the values by CWs
# create the other 3 columns
df['done as plan'] = 0
df['ideal remaining tasks'] = 0
df['actual remaining tasks'] = 0

# check if the tasks status 'closed' AND due date CW earlier or equal than the last status change CWs
for i in excel_df.index:
    if excel_df.loc[i, 'Status'] == 'Closed' and excel_df.loc[i, 'Last Status Change CWs'] <= excel_df.loc[
        i, 'due date CWs']:
        excel_df.loc[i, 'completed as planned'] = True # column 'completed as planned' shows Boolean which displays if the tasks completed
        CW_nr = excel_df.loc[i, 'due date CWs']
        df.loc[df['CWs'] == CW_nr, 'done as plan'] += 1 # column 'done as plan' shows the count of tasks completed as plan
    else:
        excel_df.loc[i, 'completed as planned'] = False

now_date = datetime.date.today()
# get the current CW (week_num) because the CW of actual remaining tasks shall not newer than the current CW
year, week_num, day_of_week = now_date.isocalendar()

for i in np.arange(0, len(df)):
    df.loc[df.index[i], 'ideal remaining tasks'] = sum(df['planned']) - sum(df['planned'][0:i + 1])
    df.loc[df.index[i], 'actual remaining tasks'] = sum(df['planned']) - sum(df['done as plan'][0:i + 1])

df2 = df.reset_index(drop=True) # Update index after sorting data-frame

fig, ax = plt.subplots()  # Create a figure containing a single axes
ax.plot(df2['CWs'], df2['ideal remaining tasks'], label='Ideal')
current_idx_cw = df2.index[df2['CWs'] == week_num][0] # draw actual remaining task till current CW
ax.plot(df2['CWs'][0:current_idx_cw+1], df2['actual remaining tasks'][0:current_idx_cw+1], label='Actual')
bars = ax.bar(df2['CWs'], df2['done as plan'], label='Completed Tasks')
ax.bar_label(bars)
plt.legend(loc='center left')

for i,j in zip(df2['CWs'],df2['ideal remaining tasks']):
    ax.annotate(str(j),xy=(i,j-3.5))

for i,j in zip(df2['CWs'][0:current_idx_cw+1],df2['actual remaining tasks'][0:current_idx_cw+1]):
    ax.annotate(str(j),xy=(i,j+3.5))

title = "Burndown Chart"+' '+planned_for_group+' '+sprint_nr
plt.xlabel('Calendar Week')
plt.ylabel('Remaining Tasks')
plt.title(title)
plt.grid()

plt.savefig(title, format="pdf", bbox_inches="tight")
plt.show()

print('Success!')