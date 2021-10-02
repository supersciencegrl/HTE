# Updated 2021-May-16 12:23
'''See params.py for version info'''

import os
import string

import PySimpleGUI as sg

import params # Parameter file including author, copyright, version, contact details, and required parameters to run workflow

''' ALL FUNCTIONS '''
def permissionerrorpopup(error, filename):
    message = f'Permission denied: {filename}'
    message = ('\n  ').join(message[i:i+64] for i in range(0, len(message), 64)) # Line breaks
    message += '\nPlease close the file to continue.'
    
    if params.showGUI:
        layout = [  [sg.Text(message)],
                    [sg.Text('', size = (3, None)), sg.Button('OK', bind_return_key = True), sg.Text('', size = (30, None))]
                    ]

        window = sg.Window(type(error).__name__, layout, keep_on_top = True, return_keyboard_events = True, right_click_menu = right)
        window.Finalize()
        window.TKroot.focus_force()

        while True:
            try:
                event, values = window.Read()
                if event in ['OK', sg.WIN_CLOSED, 'Escape:27', chr(13)]:
                    break
                elif event == 'Exit':
                    window.close()
                    raise CloseAllWindows
            except (KeyboardInterrupt, CloseAllWindows):
                closeallwindows()
                return True # Error thrown
        window.close()
    else:
        print(error)
    return False # Error not thrown

# Set Chronect dosing tray position
def setChronectDosingTray():
    if not params.Debug:
        dropdownlist = ['Tray1', 'Tray2', 'Tray3', 'Not used']
        if params.showGUI:
            ChronectDosingTray = combobox('Tray location on Chronect Quantos', dropdownlist, defaulttext = 'Tray1', size = (8, 4), tooltip = 'Rear tray: Tray1\nMiddle tray: Tray2\nFront tray: Tray3')
        else:
            ChronectDosingTray = input('Tray location on Chronect Quantos (Tray1, Tray2, Tray3, Not used): ')
    else:
        ChronectDosingTray = 'Tray1' # Default

    if ChronectDosingTray == 'Not used':
        ChronectDosingTray = None

    return ChronectDosingTray

def getRackType():
    # Determine rack type
    racktype = None
    dropdownlist = [i[0] for i in rackparameters]
    if params.showGUI:
        racktype = combobox('Rack type:', dropdownlist, defaulttext = '96')
    else:
        racktype = input('Rack type (96, 48, 24 (4-mL), 24 (8-mL): ')

    return racktype

# Create a cell in xml for the Chronect CSL file
def xmlCell(data, datatype):
    head = '\t' * 4 + '<s:Cell>\r\n'
    if data:
        value = '\t' * 5 + f'<s:Data s:Type="{datatype.capitalize()}">{data}</s:Data>\r\n'
    else:
        value = '\t' * 5 + f'<s:Data s:Type="{datatype.capitalize()}" />\r\n'
    tail = '\t' * 4 + '</s:Cell>'

    return f'{head}{value}{tail}'

def createSubstanceCSL(chemicalName, doselocations):
    # Substance parameters
    if all([not x for x in doselocations]): # if no doses will be made for this compound
        return None
    for amount in doselocations:
        if amount and not (amount.endswith(' mg') or amount.endswith(' g')): # Only mg and g amounts allowed
            return None
    ziplist = zip(all_locations_no_zeroes, doselocations)

    # xml parameters
    newrow = '\t' * 3 + '<s:Row>'
    endrow = '\t' * 3 + '</s:Row>'

    # Global parameters
    if racktype in [96, '96']:
        vialtype = '1 mL Vials'
    elif racktype in [48, '48']:
        vialtype = '2 mL Vials'
    elif racktype in [24, '24 (4-mL)']:
        vialtype = '4 mL Vials'
    elif racktype in [24, '24 (8-mL)']:
        vialtype = '8 mL Vials'
    tapduration = 2 # s
    tapintensity = 50 # %
    tolerancemode = 'ZeroPlus' # 'ZeroPlus' or 'MinusPlus'
    tolerance = 10

    substancelist = [newrow]
    
    # Create job headers
    substancelist.append(xmlCell('_', 'Number')) # Placeholder
    substancelist.append(xmlCell(r'C:\Users\Public\Documents\Chronos\Methods\Set Config.cam', 'String'))
    substancelist.append(xmlCell(vialtype, 'String'))
    substancelist.append(xmlCell(tapduration, 'Number'))
    substancelist.append(xmlCell(tapintensity, 'Number'))
    substancelist.append(xmlCell(tolerancemode, 'String'))
    substancelist.append(xmlCell('Quantos', 'String'))
    substancelist.append(xmlCell('True', 'String'))
    substancelist.append(xmlCell('True', 'String'))
    substancelist.extend(8 * [xmlCell('', 'String')],)
    substancelist.append(endrow)

    # Create each dosing event
    for (location, doseamount) in ziplist:
        location = location if location[1] != '0' else f'{location[0]}{location[2]}' # ensure NO zeroes for Chronect
        numberVialPosition = all_locations_no_zeroes.index(location) + 1
        if doseamount.endswith(' mg'):
            doseamount = doseamount[:-3]
        elif doseamount.endswith(' g'):
            doseamount = str(float(doseamount[:-2]) * 1000)
        if doseamount:
            substancelist.append(newrow)
            substancelist.append(xmlCell('_', 'Number')) # Placeholder
            substancelist.append(xmlCell(r'C:\Users\Public\Documents\Chronos\Methods\Dosing Method.cam', 'String'))
            substancelist.extend(4 * [xmlCell('', 'String')],)
            substancelist.append(xmlCell('Quantos', 'String'))
            substancelist.extend(2 * [xmlCell('', 'String')],)
            substancelist.append(xmlCell(chemicalName, 'String'))
            substancelist.append(xmlCell(ChronectDosingTray, 'String'))
            substancelist.append(xmlCell(numberVialPosition, 'Number'))
            substancelist.append(xmlCell(location, 'String')) # Must NOT have leading zero for column
            substancelist.append(xmlCell(doseamount, 'Number'))
            substancelist.append(xmlCell(tolerance, 'Number'))
            substancelist.extend(2 * [xmlCell('', 'String')],)
            substancelist.append(endrow)

    substanceCSL = ('\r\n').join(substancelist)

    return substanceCSL

def createChronectInput():
    # xml parameters
    newrow = '\t' * 3 + '<s:Row>'
    endrow = '\t' * 3 + '</s:Row>'

    # Global parameters
    chkRepeatSchedule = 'False'
    chkPrioritySchedule = 'False'
    chkOverlappedSchedule = 'True'
    
    # Spreadsheet appearance parameters
    colwidths = [22, 89, 104, 135, 142, 88, 41, 89, 89, 92, 97, 96, 126, 75, 80, 60, 56]
    CSLheaders = ['Analysis Method', 'Quantos Tray Type', 'PreDose Tap Duration [s]', 'PreDose Tap Intensity [%]',
                  'Tolerance Mode', 'Device', 'Use Front Door?', 'Use Side Doors?', 'Substance Name',
                  'Quantos Vial Tray', 'Quantos Vial Pos.', 'Quantos Vial Pos. [Axx]', 'Amount [mg]',
                  'Tolerance [%]', 'Sample ID', 'Comment']

    # Create header
    ChronectList = ['<?xml version = "1.0"?>']
    ChronectList.append('<?mso-application progid=\'Excel.Sheet\'?>')
    ChronectList.append('<s:Workbook xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:s="urn:schemas-microsoft-com:office:spreadsheet">')
    ChronectList.append('\t<s:Worksheet s:Name="grdSampleList">')
    ChronectList.append('\t\t<s:Table>')
    for colwidth in colwidths:
        ChronectList.append('\t' * 3 + f'<s:Column s:Width="{colwidth}" />')
    ChronectList.append(newrow)
    ChronectList.append(xmlCell('', 'String'))
    for header in CSLheaders:
        ChronectList.append(xmlCell(header, 'String'))
    ChronectList.append(endrow)

    for substanceCSL in CSLsubstances:
        if substanceCSL:
            ChronectList.append(substanceCSL)

    # Create tail
    ChronectList.append('\t\t</s:Table>')
    ChronectList.append('\t</s:Worksheet>')
    ChronectList.append('\t<s:Worksheet s:Name="grdSettings">')
    ChronectList.append('\t\t<s:Table>')
    ChronectList.extend(2 * ['\t\t\t<s:Column s:Width="50" />'],)
    ChronectList.append(newrow)
    ChronectList.append(xmlCell('chkRepeatSchedule', 'String'))
    ChronectList.append(xmlCell(chkRepeatSchedule, 'String'))
    ChronectList.append(endrow)
    ChronectList.append(newrow)
    ChronectList.append(xmlCell('chkPrioritySchedule', 'String'))
    ChronectList.append(xmlCell(chkPrioritySchedule, 'String'))
    ChronectList.append(endrow)
    ChronectList.append(newrow)
    ChronectList.append(xmlCell('chkOverlappedSchedule', 'String'))
    ChronectList.append(xmlCell(chkOverlappedSchedule, 'String'))
    ChronectList.append(endrow)
    ChronectList.append('\t\t</s:Table>')
    ChronectList.append('\t</s:Worksheet>')
    ChronectList.append('</s:Workbook>')

    ChronectInputWritten = False
    while not ChronectInputWritten:
        try:
            with open(ChronectInputfile, 'w', newline = '', encoding = 'utf-16') as fout:
                i = 1
                for line in ChronectList:
                    if 's:Type="Number">_</s:Data>' in line:
                        linelist = line.split('s:Type="Number">_</s:Data>')
                        newline = ''
                        for section in linelist[:-1]:
                            newline = newline + section + f's:Type="Number">{i}</s:Data>'
                            i += 1
                        newline = newline + linelist[-1]
                        line = ('').join(newline)
                    fout.write(line + '\r\n')
        except PermissionError as error:
            filename = os.path.join(mydir, ChronectInputfile)
            permissionerrorpopup(error, filename)
        else:
            ChronectInputWritten = True
        
    addfinalmessage(f'{ChronectInputfile} Chronect inputfile written.')

# GUI functions
def combobox(question, dropdownlist, **kwargs):
    if params.showGUI:
        # kwargs
        defaulttext = ''
        tooltip = None
        if 'defaulttext' in kwargs:
            defaulttext = kwargs['defaulttext']
        if 'size' in kwargs:
            comboboxsize = kwargs['size']
        else:
            comboboxsize = (5, 4)
        if 'tooltip' in kwargs:
            tooltip = kwargs['tooltip']

        layout = [  [sg.Text(question.strip(), size = (30, None), tooltip = tooltip)],
                    [sg.Sizer(30), sg.Combo(dropdownlist, default_value = '_', key = 'result', size = comboboxsize, enable_events = True)],
                    [sg.Button('OK', bind_return_key = True), sg.Text('', size = (23, None), right_click_menu= right)]
                    ]

        window = sg.Window('Choose parameters', layout, keep_on_top = True, return_keyboard_events = True, right_click_menu = right, use_default_focus = False)
        window.Finalize()
        window.TKroot.focus_force()
        window.Element('result').SetFocus()
        window['result'].update(defaulttext)

        i = 0
        while True:
            try:
                event, values = window.Read()

                ''' You may need to add a delay or other function: loading the dialog by pressing Enter too slowly will trigger event == chr(13) '''
                if event in ['OK', chr(13), sg.WIN_CLOSED]: # chr(13) == pressing Enter with focus on combobox
                    if values['result']:
                        break
                elif event == 'Exit':
                    window.close()
                    exit()
            except KeyboardInterrupt:
                window.close()
        window.close()
        result = values['result']
    return result

def addfinalmessage(message):
    global finalmessage
    
    print(message)
    if not message.strip():
        finalmessage.append(message.strip()) # Only strip if this leaves a non-zero length string
    else:
        finalmessage.append(message)

def finalpopup(finalmessage, title):
    message = ('\n').join(finalmessage)
    
    if params.showGUI:
        buttoncolumn = [    [sg.Button('Open Chronect input', key = 'Quantos')],
                            [sg.Button('Close', bind_return_key = True)]
                            ]
        
        layout = [  [sg.Text(message), sg.Column(buttoncolumn)]
                    ]

        window = sg.Window(title, layout, keep_on_top = True, return_keyboard_events = True, right_click_menu = right)
        window.Finalize()
        window.TKroot.focus_force()

        while True:
            try:
                event, values = window.Read()

                ''' You may need to add a delay or other function: loading the dialog by pressing Enter too slowly will trigger event == chr(13) '''
                if event in ['Close', sg.WIN_CLOSED, 'Escape:27', 'Exit', chr(13)]:
                    break
                elif event == 'Quantos':
                    print('Opening Chronect Quantos inputfile.')
                    openchronectfile()
            except KeyboardInterrupt:
                break
        window.close()
    else:
        print(error)

# Open output files
def openchronectfile():
    os.system(f'start excel "{os.path.join(mydir, ChronectInputfile)}"')

# For faster debugging
def killexcel():
    os.system('taskkill /f /im excel.exe')

sg.theme('Green')
right = ['right', ['Exit']]
finalmessage = []

pydir = params.pydir
mydir = pydir
os.chdir(mydir)

# Get rack type
''' All possible rack types should be in your rackparameters.
    Format: name(str), rows(int), cols(int)'''
rackparameters = [['96', 8, 12], ['48', 6, 8], ['24 (4-mL)', 4, 6], ['24 (8-mL)', 4, 6]]
racktype = getRackType()

platerows = list(string.ascii_uppercase)[:8] # Letters A-H
platecols = [str(x + 1) for x in range(24)]

for p in rackparameters:
    if racktype == p[0]:
        all_locations_no_zeroes = [(row + col) for row in platerows[:p[1]] for col in platecols[:p[2]]]
        break

# Get Chronect dosing tray location
ChronectDosingTray = setChronectDosingTray()

''' Here is the example input for the createSubstanceCSL function. You could automatically generate this from an eLN table, or user Excel input, for example '''
examplesubstances = [   {'Name': 't-BuBrettPhos', 'DoseLocations': ['', '1.0 mg', '', '1.0 mg', '', '1.0 mg', '', '1.0 mg', '', '1.0 mg', '', '1.0 mg', '', '1.0 mg', '', '1.0 mg', '', '1.0 mg', '', '1.0 mg', '', '', '', '', '', '', '', '']},
                        {'Name': 'CataCXium A', 'DoseLocations': ['2.1 mg', '2.1 mg', '2.1 mg', '2.1 mg', '2.1 mg', '2.1 mg', '2.1 mg', '2.1 mg', '0.5 mg', '', '0.5 mg', '', '0.5 mg', '', '0.5 mg', '', '', '', '', '', '', '', '', '']}
                        ]

CSLsubstances = []
for substance in examplesubstances:
    substanceCSL = createSubstanceCSL(substance['Name'], substance['DoseLocations'])
    CSLsubstances.append(substanceCSL)

# Write Chronect inputfile
if ChronectDosingTray:
    ChronectInputfile = f'inputfile.csl' ### Default name
    createChronectInput()

if finalmessage:
    finalpopup(finalmessage, 'Successful!')
