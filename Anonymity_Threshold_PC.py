import pandas as pd
import time
import datetime
import tkinter as tk
import webbrowser
from tkinter import filedialog,messagebox, ttk
import os

start_time = time.time()

########################################## Execute the script ##########################################
def run_mix():
    #Get file paths from user search
    ExportUnits_file_path = file_entry_data1.get()
    Response_file_path = file_entry_data2.get()
    Participants_file_path = file_entry_data3.get()
    
    #Get Employee and Manager IDs
    Empl_Man_IDs  = Empl_Man_IDs_Columns.get().split(',')

    #getting column Names
    #columnNames = ['Unique Identifier', 'ManagerID']

    #Employee ID
    EmpID = Empl_Man_IDs[0]

    #response file - Employee ID
    if EmpID == 'Unique Identifier':
        ResponseEmployeeID = 'Participant Unique Identifier'
    else:
        ResponseEmployeeID = EmpID

    #Manager ID
    ManID_Spaces = Empl_Man_IDs[1]
    ManID = ManID_Spaces.lstrip()

    # obtaining the anonymity threshold value
    anonymity_threshold_value = int(anonymity_threshold_entry.get())

    #error handling
    if not ExportUnits_file_path or not Response_file_path or not Participants_file_path or not Empl_Man_IDs_Columns:
        messagebox.showerror("Error", "Please provide all file paths.")
        return
    
    ########################### read files ############################################################
    # processing Participant File
    Participants = pd.read_csv(f'{Participants_file_path}', dtype='str')

    # processing Export Unit file
    ExportUnitsDf = pd.read_csv(f'{ExportUnits_file_path}', dtype='str')

    # processing D&A file
    Responses = pd.read_csv(f'{Response_file_path}', dtype='str')


    ############################ Participants File ################################################################
    # obtaining the employee and manager IDs from the Participants file
    Participants_IDs = Participants[[EmpID, ManID, 'Respondent']]
    Participants_Metadata = Participants[['First Name', 'Last Name', 'Email', EmpID]]

    # removing blank employee ID participants
    Participants_NoNulls = Participants_IDs.dropna(subset=[EmpID])

    # replacing '' with 0 for participants without Manager ID
    #Participants_ManagerID_NotNull = Participants_NoNulls.fillna(0)

    # filtering non respondent participants
    Participants_IDs_Filtered = Participants_NoNulls.loc[Participants_IDs['Respondent'] == 'true']

    # setting Employee ID and Manager ID to be strings
    Participants_InvitedCount = Participants_IDs_Filtered[[EmpID, ManID]].astype(str)
    Participants_Path = Participants_NoNulls[[EmpID, ManID]].astype(str)

############################ Export Units File ################################################################
    exportUnits_IDS = ExportUnitsDf[[ManID]]

# removing blank employee ID participants
    exportUnits_NoNulls = exportUnits_IDS.dropna(subset=ManID)

# removing "No Manager" rows from Manager ID column
    exportUnits_IDs_Final = exportUnits_NoNulls[~exportUnits_NoNulls[ManID].str.contains('NoManager')]

# setting Manager ID to be strings
    exportUnits_Final = exportUnits_IDs_Final[[ManID]].astype(str)

############################ Response File ################################################################
    response_IDS = Responses[[ResponseEmployeeID]].iloc[2:]

# removing blank employee id responses
    response_NoNulls = response_IDS.dropna(subset=[ResponseEmployeeID])

#replacing ".0" with nothing
    response_NoDecimals = response_NoNulls.replace('.0', '')

# converting employee ID to be string
    response_Final = response_NoDecimals[[ResponseEmployeeID]].astype(str)

# adding response column
    response_Final['Response'] = 'Yes'

    ############################ Path generation ################################################################
    # assigning values of Employee ID and Manager ID columns into a list
    hierarchy = Participants_Path[[EmpID, ManID]].values.tolist()

    # obtaining the paths for all the employees
    manager = set()
    employee = {}

    for e,m in hierarchy:
        manager.add(m)
        employee[e] = m

    # recursively determine parents until child has no parent, appending parents at the beginning of the list
    def ancestors(p):
        return (ancestors(employee[p]) if p in employee else []) + [p]

    # creating the path for each employee
    path = {}
    for k in (set(employee.keys())):
        # creating one path for each employee
        pathstr = '/'.join(ancestors(k))
        # removing 0/ from the paths
        #pathStrClean = pathstr.replace(pathstr[:2], '',1)
        path[k] = pathstr

    # assigning the paths to a data frame
    pathDf = pd.DataFrame(list(path.items()))

    # renaming the "Path" column
    pathDf.rename(columns={1: 'Path'}, inplace=True)

    # joining the path df with the participants df
    ParticipantsWithPath = pd.merge(Participants_InvitedCount, pathDf, left_on=EmpID, right_on=0, how= 'left').drop(0, axis=1)

    # joining the response df with the participants df
    ParticipantsWithResponses = pd.merge(ParticipantsWithPath, response_Final, left_on=EmpID, right_on=ResponseEmployeeID, how= 'left').drop(ResponseEmployeeID, axis=1)
    ParticipantsWithResponses['Response'] = ParticipantsWithResponses[['Response']].fillna('No')

    # joining the path df with the Export Units df
    ExportUnitsWithPath = pd.merge(exportUnits_Final, pathDf, left_on=ManID, right_on=0, how= 'left').drop(0, axis=1)

    ############################ Calculating the Invited and Response Counts ################################################################
    # calculating the expected count - No Response
    # converting the path columns into a list
    ParticipantsPathNoReponse = ParticipantsWithResponses.loc[ParticipantsWithResponses['Response'] == 'No']
    ParticipantsPathNRList = ParticipantsPathNoReponse[['Path']].sort_values('Path').values.tolist()
    # converting the list into a string
    ParticipantsPathStrNR = ' '.join([str(elem) for elem in ParticipantsPathNRList])

    # filtering participants who responded the survey
    ParticipantsPathR = ParticipantsWithResponses.loc[ParticipantsWithResponses['Response'] == 'Yes']
    # converting the path columns into a list
    ParticipantsPathRList = ParticipantsPathR[['Path']].values.tolist()

    # converting the list into a string
    ParticipantsPathStrR = ' '.join([str(elem) for elem in ParticipantsPathRList])

    expected_count_NR_dic = {}
    response_count_dic = {}
    count = 0
    for index, row in ExportUnitsWithPath.iterrows():
        correctedPath = row['Path'] + '/'
        # calculating the expected count
        expectedCountNR = ParticipantsPathStrNR.count(correctedPath)
        expected_count_NR_dic[row[ManID]] = expectedCountNR
        # calculating the response count
        responseCount = ParticipantsPathStrR.count(correctedPath)
        response_count_dic[row[ManID]] = responseCount

        print(count, expectedCountNR, responseCount)
        count += 1

    # assigning the paths to a data frame
    expectedCountDf = pd.DataFrame(list(expected_count_NR_dic.items()))
    responseCountDf = pd.DataFrame(list(response_count_dic.items()))

    # renaming the "ExpectedCount" column
    expectedCountDf.rename(columns={1: 'InvitedCountNR'}, inplace=True)
    responseCountDf.rename(columns={1: 'ResponseCount'}, inplace=True)

    # joining the path df with the participants df
    ExportUnitsExpectedCountNR = pd.merge(ExportUnitsWithPath, expectedCountDf, left_on=ManID, right_on=0, how= 'left').drop(0, axis=1)
    dfAnonymityThreshold = pd.merge(ExportUnitsExpectedCountNR, responseCountDf, left_on=ManID, right_on=0, how= 'left').drop(0, axis=1)

    # adding the No response values with the response count, to get the invited count
    dfAnonymityThreshold['InvitedCount'] = dfAnonymityThreshold['InvitedCountNR'] + dfAnonymityThreshold['ResponseCount']

    #dfAnonymityThreshold.sort_values(ManID).to_csv('PathReport.csv')

    # deleting invited count for non respondents
    del dfAnonymityThreshold['InvitedCountNR']

    # calculating the anonymity threshold
    dfAnonymityThreshold['Anonymity Threshold'] = [True if x >= anonymity_threshold_value else False for x in dfAnonymityThreshold['ResponseCount']]

    # removing "nan/" from Path
    dfAnonymityThreshold['Path'] = dfAnonymityThreshold['Path'].str[4:]

    # adding the metadata fields to the report
    dfAnonymityThresholdReport = pd.merge(dfAnonymityThreshold, Participants_Metadata, left_on=ManID, right_on=EmpID, how= 'left').drop(EmpID, axis=1)
    '''SAVE DOCUMENT'''

    # Create Folder to save the report
    new_filename_to_save = Participants_file_path.split('/')
    
    # extract the file name to delete ".csv" string
    final_filename = new_filename_to_save[-1]

    # extract the number of characters to be deleted from the original path and assign to the new folder
    len_user_file = len(final_filename) + 1

    # delete the length of the user file to create the original path
    new_path_to_save = Participants_file_path[0:-len_user_file]

    # create folder to save the report
    reportPath = new_path_to_save + '/Anonymity_Report'
    if not os.path.exists(reportPath):
        os.makedirs(reportPath)

    #Save new DF as CSV
    def save_doc():
        today = datetime.datetime.today()
        timestamp_str = today.strftime("%Y%m%d_%H%M")
        dfAnonymityThresholdReport[['First Name', 'Last Name', 'Email', ManID, 'Path', 'InvitedCount', 'ResponseCount', 'Anonymity Threshold']].sort_values(ManID).to_csv(f'{new_path_to_save}/Anonymity_Report/AnonymityReport{timestamp_str}.csv',index=False)
        return True
    
    #Call Save_doc function
    save_status = save_doc()
    end_time = time.time()        
    print("Execution time in seconds:", end_time-start_time)  
    
    if save_status:
        messagebox.showerror("Success", "Doc was saved.")
    else:
        messagebox.showerror("Failed", "Doc was not saved due to an error.") 

#--------------------------User Interface -----------------------------------------
    
def browse_files1():
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    file_entry_data1.delete(0, tk.END)
    file_entry_data1.insert(0, file_path)

def browse_files2():
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    file_entry_data2.delete(0, tk.END)
    file_entry_data2.insert(0, file_path)
    
def browse_files3():
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    file_entry_data3.delete(0, tk.END)
    file_entry_data3.insert(0, file_path)
    
# Set up the main application window
root = tk.Tk()
root.title("Manager Anonymity Threshold Checkr")

# File selection
file_label_data1 = tk.Label(root, text="Export Unit Ids File path")
file_label_data1.grid(row=0, column=0, padx=20, pady=20, sticky=tk.W)
file_entry_data1 = tk.Entry(root, width=30)
file_entry_data1.grid(row=0, column=1, padx=20, pady=20)
browse_button_data1 = tk.Button(root, text="Browse",command=browse_files1)
browse_button_data1.grid(row=0, column=2, padx=20, pady=20)

file_label_data2 = tk.Label(root, text="Response File path")
file_label_data2.grid(row=1, column=0, padx=20, pady=20, sticky=tk.W)
file_entry_data2 = tk.Entry(root, width=30)
file_entry_data2.grid(row=1, column=1, padx=20, pady=20)
browse_button_data2 = tk.Button(root, text="Browse",command=browse_files2)
browse_button_data2.grid(row=1, column=2, padx=20, pady=20)

file_label_data3 = tk.Label(root, text="Participant File path")
file_label_data3.grid(row=2, column=0, padx=20, pady=20, sticky=tk.W)
file_entry_data3 = tk.Entry(root, width=30)
file_entry_data3.grid(row=2, column=1, padx=20, pady=20)
browse_button_data3 = tk.Button(root, text="Browse",command=browse_files3)
browse_button_data3.grid(row=2, column=2, padx=20, pady=20)

# Level based column names input
Empl_Man_IDs_Columns = tk.Label(root, text="Enter the names of the columns of Employee and Manager ID (comma-separated):")
Empl_Man_IDs_Columns.grid(row=3, column=0, padx=20, pady=20, sticky=tk.W)
Empl_Man_IDs_Columns = tk.Entry(root, width=30)
Empl_Man_IDs_Columns.grid(row=3, column=1, padx=0, pady=20, columnspan=2)

#Set anonymity Threshold value
anonymity_threshold_label = tk.Label(root, text="Set anonymity threshold value")
anonymity_threshold_label.grid(row=4, column=0, padx=20, pady=20, sticky=tk.W)
anonymity_threshold_entry = tk.Entry(root, width=30)
anonymity_threshold_entry.grid(row=4, column=1, padx=0, pady=20, columnspan=2)


unsupported_chars_list_frame = tk.LabelFrame(root, text = "HOW TO USE:")
unsupported_chars_list_frame.grid(row=5, column=0,columnspan=10, padx=20, pady=30, sticky=tk.W)

unsupported_chars_list_1=tk.Label(unsupported_chars_list_frame, text='1.Select the files to perform the analysis. ')
unsupported_chars_list_1.grid(row=1, column=1, padx=20, pady=20, sticky=tk.W)
unsupported_chars_list_2=tk.Label(unsupported_chars_list_frame, text='2.Enter the name of the columns for Employee and Manager ID. ')
unsupported_chars_list_2.grid(row=2, column=1, padx=20, pady=20, sticky=tk.W)
unsupported_chars_list_3=tk.Label(unsupported_chars_list_frame, text='3.Set the Anonymity Threshold value')
unsupported_chars_list_3.grid(row=3, column=1, padx=20, pady=20, sticky=tk.W)
unsupported_chars_list_4=tk.Label(unsupported_chars_list_frame, text='4.Click "Get Anonymity Threshold!"')
unsupported_chars_list_4.grid(row=4, column=1, padx=20, pady=20, sticky=tk.W)

# Run mix button
run_button = tk.Button(root, text="Get Anonymity Threshold!", command=run_mix)
run_button.grid(row=6, column=0, columnspan=3, padx=20, pady=20)

new=1
url = "https://coda.io/d/QA-Automation_diYJsprOr4k/Managers-Anonymity-Threshold_suB1Q#_luXuW"

def openweb():
    webbrowser.open(url,new=new)

feedback_button= tk.Button(root, text ="Feedback / New ideas", command=openweb)
feedback_button.grid(row=5, column=1, columnspan=3, padx=20, pady=20)

# Start the application
root.mainloop()






