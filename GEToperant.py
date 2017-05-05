### This program will collected Med PC data files and match them to a data profile for saving to an excel workbook.
### Preconditions: All data files must have the same structure and fit the same profile.
### It is recommended to open either multiple files with single subjects or one file with multiple subjects.
### Where an MSN has not used the Y2KCOMPLIANT command, the data was collected in the 21st century.

import xlrd
import xlsxwriter
import re
import itertools

def GEToperant(GETprofile, MPCdatafiles, outputfile):
    '''
    GEToperant takes three arguments:
    GETprofile, which must be an Excel file
    MPCdatafiles, which must be a list of one or more Med-PC data files
    outputfile, which must be an Excel file

    GEToperant will read the data from the MPCdatafiles and will
    output the headers and the data described in GETprofile. It will
    save this in the Excel file specified by outputfile.

    Preconditions: All sessions in each file must have the same data.
    If arrays or variables are present in some of the MPC data but not others
    then this will cause the output data to be sorted incorrectly.
    The profile must be designed according to the instructions in the GUI.
    '''

    ### This first part will read the data profile and develop a series of lists
    profile_xl = xlrd.open_workbook(GETprofile)
    profile_xl_sheets = profile_xl.sheet_names()
    profilesheet = profile_xl.sheet_by_name(profile_xl_sheets[0])

    Label = list()
    LabelStartValue = list()
    LabelIncrement = list()
    ArrayVar = list()
    StartElement = list()
    ArrayIncrement = list()
    StopElement = list()

    for r in range(1,max(range(profilesheet.nrows))):
        cell0 = profilesheet.cell(r,0)
        Label.append(str(cell0).split("\'")[1])
        
        cell1 = profilesheet.cell(r,1)
        if 'empty' in str(cell1):
            LabelStartValue.append(None)
        elif 'number' in str(cell1):
            LabelStartValue.append(int(float(str(cell1).split(":")[1])))

        cell2 = profilesheet.cell(r,2)
        if 'empty' in str(cell2):
            LabelIncrement.append(None)
        elif 'number' in str(cell2):
            LabelIncrement.append(int(float(str(cell2).split(":")[1])))

        cell3 = profilesheet.cell(r,3)
        ArrayVar.append(str(cell3).split("\'")[1])

        cell4 = profilesheet.cell(r,4)
        StartElement.append(int(float(str(cell4).split(":")[1])))

        cell5 = profilesheet.cell(r,5)
        if 'empty' in str(cell5):
            ArrayIncrement.append(None)
        elif 'number' in str(cell5):
            ArrayIncrement.append(int(float(str(cell5).split(":")[1])))

        cell6 = profilesheet.cell(r,6)
        if 'empty' in str(cell6) or 'text' in str(cell6):
            StopElement.append(None)
        elif 'number' in str(cell6):
            StopElement.append(int(float(str(cell6).split(":")[1])))

    ### The relevant fields in the Med-PC file are then defined as a series of lists
    Filenames = list()
    Startdate = list()
    Enddate = list()
    Subject = list()
    Experiment = list()
    Group = list()
    Box = list()
    Starttime = list()
    Endtime = list()
    MSN = list()
    A = list()
    B = list()
    C = list()
    D = list()
    E = list()
    F = list()
    G = list()
    H = list()
    I = list()
    J = list()
    K = list()
    L = list()
    M = list()
    N = list()
    O = list()
    P = list()
    Q = list()
    R = list()
    S = list()
    T = list()
    U = list()
    V = list()
    W = list()
    X = list()
    Y = list()
    Z = list()
    Comments = list()

    ### Values will hold the numbers for each array so they can be collected and flattened
    values = list()
    currentarray = ''

    MPC_filelist = list(MPCdatafiles)
    MPC_file = list()
    for i in MPC_filelist:
        MPC_file.append(open(i, 'r').readlines())

    ### Begin the for loop that will loop over the data and collect everything into MPC_file
    for i in MPC_file:
        for line in i:
            # Begin by collecting the headers
            # Collect the file names
            if 'File' in line:
                path = line[6:len(line)-1]
                Filenames.append(path)
            # Collect the start and end dates in ISO 8601 format, correcting for a lack of Y2KCOMPLIANT.
            elif 'Start Date' in line:
                if len(line) < 22:
                    Startdate.append("20"+line[18:len(line)-1]+"-"+line[12:14]+"-"+line[15:17])
                else:
                    Startdate.append(line[18:len(line)-1]+"-"+line[12:14]+"-"+line[15:17])
                if len(Startdate) > len(Filenames):
                    Filenames.append(None)
            elif 'End Date' in line:
                if len(line) < 20:
                    Enddate.append("20"+line[16:len(line)-1]+"-"+line[10:12]+"-"+line[13:15])
                else:
                    Enddate.append(line[16:len(line)-1]+"-"+line[10:12]+"-"+line[13:15])
            # Similarly, collect subject, experiment, group, box, start time, end time and program name
            elif 'Subject' in line:
                Subject.append(line[9:len(line)-1])
            elif 'Experiment' in line:
                Experiment.append(line[12:len(line)-1])
            elif 'Group' in line:
                Group.append(line[7:len(line)-1])
            elif 'Box' in line:
                Box.append(line[5:len(line)-1])
            elif 'Start Time' in line:
                if line[12] == ' ':
                    Starttime.append(line[13:len(line)-1])
                else:
                    Starttime.append(line[12:len(line)-1])
            elif 'End Time' in line:
                if line[10] == ' ':
                    Endtime.append(line[11:len(line)-1])
                else:
                    Endtime.append(line[10:len(line)-1])
            elif 'MSN' in line:
                MSN.append(line[5:len(line)-1])
            # Check for an array header, if it is present, check if values have been entered into
            # a previous data array. If there are previous data values, flatten the data array and dump them.
            elif len(line) > 1:
                part_check = re.search(r'\D:', line)
                if part_check != None:
                    if len(values) > 0:
                        values = list(itertools.chain.from_iterable(values))
                        eval(currentarray).append(values)
                        values = list()
                    ### here we should check for whether the letter has been printed as just a variable.
                    part_checkb = re.search(r'\d', line)
                    if part_checkb != None:
                        currentarray = line[0]
                        values.append(line.split()[1])
                    ### then we should set the beginning of a new array.
                    else:
                        currentarray = line[0]
                ### this part should then collect data into a new array
                else:
                    values.append(line.split()[1:])
                if re.search(r'[\\]', line) != None:
                    Comments.append(line[1:len(line)-1])
            elif len(line) < 1 and len(values) > 0:
                values = list(itertools.chain.from_iterable(values))
                eval(currentarray).append(values)
                values = list()
            elif len(Startdate) > len(Comments):
                    Comments.append(None)

    ### This final part will begin writing the data to the Excel file.
    output = xlsxwriter.Workbook(outputfile)
    output.set_properties({
        'title': 'Med-PC Data',
        'subject': 'Animal behaviour',
        'category': 'Raw data',
        'comments': 'Extracted using GEToperant, a Python program using xlrd and xlsxwriter. https://www.github.com/SKhoo'
        })

    mainsheet = output.add_worksheet('GEToperant output')

    ### Write the headers
    mainsheet.set_column('A:A', 15)
    mainsheet.write(0, 0, 'Filename')
    for i in range(len(Filenames)):
        mainsheet.write(0, i+1, Filenames[i])

    mainsheet.write(1, 0, 'Start Date')
    for i in range(len(Startdate)):
        mainsheet.write(1, i+1, Startdate[i])

    mainsheet.write(2, 0, 'End Date')
    for i in range(len(Enddate)):
        mainsheet.write(2, i+1, Enddate[i])

    mainsheet.write(3, 0, 'Subject')
    for i in range(len(Subject)):
        mainsheet.write(3, i+1, Subject[i])

    mainsheet.write(4, 0, 'Experiment')
    for i in range(len(Subject)):
        mainsheet.write(4, i+1, Experiment[i])

    mainsheet.write(5, 0, 'Group')
    for i in range(len(Group)):
        mainsheet.write(5, i+1, Group[i])

    mainsheet.write(6, 0, 'Box')
    for i in range(len(Box)):
        mainsheet.write(6, i+1, float(Box[i]))

    mainsheet.write(7, 0, 'Start Time')
    for i in range(len(Starttime)):
        mainsheet.write(7, i+1, Starttime[i])

    mainsheet.write(8, 0, 'End Time')
    for i in range(len(Endtime)):
        mainsheet.write(8, i+1, Endtime[i])

    mainsheet.write(9, 0, 'MSN')
    for i in range(len(MSN)):
        mainsheet.write(9, i+1, MSN[i])

    lastrow = 9

    for i in range(len(Label)):
        ### This function will loop over the profile. For each label it will check if it is
        ### 1. A single element extraction
        ### 2. A partial array extraction
        ### 3. A full array extraction
        if ArrayIncrement[i] < 1:
            # Single element extraction takes only the label
            lastrow = lastrow + 1
            mainsheet.write(lastrow, 0, Label[i])
            if 'comment' in ArrayVar[i].lower():
                for k in range(len(Subject)):
                    if k < len(Comments):
                        mainsheet.write(lastrow, k+1, Comments[k])
                    else:
                        mainsheet.write(lastrow, k+1, None)
            else:
                for k in range(len(Subject)):
                    mainsheet.write(lastrow, k+1, float(eval(ArrayVar[i])[k][StartElement[i]]))
        elif ArrayIncrement[i] > 0:
            if StopElement[i] == None or isinstance(StopElement[i], str):
                steps = range(StartElement[i], len(max(eval(ArrayVar[i]), key = len)), ArrayIncrement[i])
            elif StopElement[i] > StartElement[i]:
                steps = range(StartElement[i], StopElement[i] + 1, ArrayIncrement[i])
            for x in steps:
                lastrow = lastrow + 1
                for k in range(len(Subject)):
                    if LabelIncrement[i] != None and LabelIncrement[i] > 0:
                        mainsheet.write(lastrow, 0, Label[i] + ' ' + str(LabelStartValue[i] + x * LabelIncrement[i]))
                    else:
                        mainsheet.write(lastrow, 0, Label[i])
                    if x < len(eval(ArrayVar[i])[k]):
                        mainsheet.write(lastrow, k+1, float(eval(ArrayVar[i])[k][x]))
                    else:
                        mainsheet.write(lastrow, k+1, None)
