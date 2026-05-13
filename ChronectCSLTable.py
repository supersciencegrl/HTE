'''See params.py for version info'''
import os
import string
import textwrap

import PySimpleGUI as sg

import params # Parameter file including author, copyright, version, contact details, and required parameters to run workflow

''' ALL FUNCTIONS '''
def permission_error_popup(error, filename):
    message = f'Permission denied: {filename}. Please close the file to continue.'
    wrapped = textwrap.fill(message, width = 64, break_long_words = True, break_on_hyphens=True)
    
    if not params.SHOW_GUI:
        print(error, wrapped)
        return None
    
    layout = [  [sg.Text(wrapped)],
                [sg.Text('', size = (3, None)), sg.Button('OK', bind_return_key = True), sg.Text('', size = (30, None))]
                ]

    window = sg.Window(type(error).__name__, layout, keep_on_top = True, return_keyboard_events = True, right_click_menu = right)
    window.finalize()
    window.TKroot.focus_force()

    while True:
        try:
            event, _ = window.read()
            if event in ['OK', sg.WIN_CLOSED, 'Escape:27', '\r', chr(13)]:
                break
            elif event == 'Exit':
                window.close()
                raise CloseAllWindows
        except (KeyboardInterrupt, CloseAllWindows):
            close_all_windows()
        finally:
            try:
                window.close()
            except Exception:
                pass

# Set Chronect dosing tray position
def set_chronect_dosing_tray():
    if not params.DEBUG:
        dropdown_list = ['Tray1', 'Tray2', 'Tray3', 'Not used']
        tooltip = 'Rear tray: Tray1\nMiddle tray: Tray2\nFront tray: Tray3'
        if params.SHOW_GUI:
            chronect_dosing_tray = combobox('Tray location on Chronect Quantos', 
                                            dropdown_list, 
                                            default_text = 'Tray1', 
                                            size = (8, 4), 
                                            tooltip = tooltip
                                            )
        else:
            prompt = f'Tray location on Chronect Quantos ({(', ').join(dropdown_list)}): '
            chronect_dosing_tray = input(prompt)
    else:
        chronect_dosing_tray = 'Tray1' # Default

    if chronect_dosing_tray == 'Not used':
        chronect_dosing_tray = None

    return chronect_dosing_tray

def get_rack_type():
    # Determine rack type
    rack_type = None
    dropdown_list = [i[0] for i in rack_parameters]
    if params.SHOW_GUI:
        rack_type = combobox('Rack type:', dropdown_list, default_text = '96')
    else:
        rack_type = input('Rack type (96, 48, 24 (4-mL), 24 (8-mL): ')

    return rack_type

# Create a cell in xml for the Chronect CSL file
def create_xml_cell(data: str|int|float, data_type: str):
    data_type = data_type.capitalize()

    head = '\t' * 4 + '<s:Cell>\r\n'
    if data:
        value = '\t' * 5 + f'<s:Data s:Type="{data_type}">{data}</s:Data>\r\n'
    else:
        value = '\t' * 5 + f'<s:Data s:Type="{data_type}" />\r\n'
    tail = '\t' * 4 + '</s:Cell>'

    return f'{head}{value}{tail}'

def create_substance_csl(chemical_name: str, dose_locations: list[str]):
    # Substance parameters
    if all(not x for x in dose_locations): # if no doses will be made for this compound
        return None
    for amount in dose_locations:
        if amount and not (amount.endswith(' mg') or amount.endswith(' g')): # Only mg and g amounts allowed
            return None
    zip_list = zip(all_locations_no_zeroes, dose_locations)

    # xml parameters
    new_row = '\t' * 3 + '<s:Row>'
    end_row = '\t' * 3 + '</s:Row>'

    # Global parameters
    if rack_type in [96, '96']:
        vial_type = '1 mL Vials'
    elif rack_type in [48, '48']:
        vial_type = '2 mL Vials'
    elif rack_type in [24, '24 (4-mL)']:
        vial_type = '4 mL Vials'
    elif rack_type == '24 (8-mL)':
        vial_type = '8 mL Vials'
    tap_duration = 2 # s
    tap_intensity = 50 # %
    tolerance_mode = 'ZeroPlus' # 'ZeroPlus' or 'MinusPlus'; generally recommend ZeroPlus
    tolerance = 10

    substance_list = [new_row]
    
    # Create job headers
    substance_list.append(create_xml_cell('_', 'Number')) # Placeholder
    substance_list.append(create_xml_cell(r'C:\Users\Public\Documents\Chronos\Methods\Set Config.cam', 'String'))
    substance_list.append(create_xml_cell(vial_type, 'String'))
    substance_list.append(create_xml_cell(tap_duration, 'Number'))
    substance_list.append(create_xml_cell(tap_intensity, 'Number'))
    substance_list.append(create_xml_cell(tolerance_mode, 'String'))
    substance_list.append(create_xml_cell('Quantos', 'String'))
    substance_list.append(create_xml_cell('True', 'String'))
    substance_list.append(create_xml_cell('True', 'String'))
    substance_list.extend(8 * [create_xml_cell('', 'String')],)
    substance_list.append(end_row)

    # Create each dosing event
    for (location, dose_amount) in zip_list:
        location = location if location[1] != '0' else f'{location[0]}{location[2]}' # ensure NO zeroes for Chronect
        number_vial_position = all_locations_no_zeroes.index(location) + 1
        if dose_amount.endswith(' mg'):
            dose_amount = dose_amount[:-3]
        elif dose_amount.endswith(' g'):
            dose_amount = str(float(dose_amount[:-2]) * 1000)
        if dose_amount:
            substance_list.append(new_row)
            substance_list.append(create_xml_cell('_', 'Number')) # Placeholder
            substance_list.append(create_xml_cell(r'C:\Users\Public\Documents\Chronos\Methods\Dosing Method.cam', 'String'))
            substance_list.extend(4 * [create_xml_cell('', 'String')],)
            substance_list.append(create_xml_cell('Quantos', 'String'))
            substance_list.extend(2 * [create_xml_cell('', 'String')],)
            substance_list.append(create_xml_cell(chemical_name, 'String'))
            substance_list.append(create_xml_cell(chronect_dosing_tray, 'String'))
            substance_list.append(create_xml_cell(number_vial_position, 'Number'))
            substance_list.append(create_xml_cell(location, 'String')) # Must NOT have leading zero for column
            substance_list.append(create_xml_cell(dose_amount, 'Number'))
            substance_list.append(create_xml_cell(tolerance, 'Number'))
            substance_list.extend(2 * [create_xml_cell('', 'String')],)
            substance_list.append(end_row)

    substance_csl = ('\r\n').join(substance_list)

    return substance_csl

def create_chronect_input():
    # xml parameters
    new_row = '\t' * 3 + '<s:Row>'
    end_row = '\t' * 3 + '</s:Row>'

    # Global parameters
    chk_repeat_schedule = 'False'
    chk_priority_schedule = 'False'
    chk_overlapped_schedule = 'True'
    
    # Spreadsheet appearance parameters
    col_widths = [22, 89, 104, 135, 142, 88, 41, 89, 89, 92, 97, 96, 126, 75, 80, 60, 56]
    csl_headers = ['Analysis Method', 'Quantos Tray Type', 'PreDose Tap Duration [s]', 'PreDose Tap Intensity [%]',
                  'Tolerance Mode', 'Device', 'Use Front Door?', 'Use Side Doors?', 'Substance Name',
                  'Quantos Vial Tray', 'Quantos Vial Pos.', 'Quantos Vial Pos. [Axx]', 'Amount [mg]',
                  'Tolerance [%]', 'Sample ID', 'Comment']

    # Create header
    chronect_list = ['<?xml version = "1.0"?>']
    chronect_list.append('<?mso-application progid=\'Excel.Sheet\'?>')
    chronect_list.append('<s:Workbook xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:s="urn:schemas-microsoft-com:office:spreadsheet">')
    chronect_list.append('\t<s:Worksheet s:Name="grdSampleList">')
    chronect_list.append('\t\t<s:Table>')
    for col_width in col_widths:
        chronect_list.append('\t' * 3 + f'<s:Column s:Width="{col_width}" />')
    chronect_list.append(new_row)
    chronect_list.append(create_xml_cell('', 'String'))
    for header in csl_headers:
        chronect_list.append(create_xml_cell(header, 'String'))
    chronect_list.append(end_row)

    for substance_csl in csl_substances:
        if substance_csl:
            chronect_list.append(substance_csl)

    # Create tail
    chronect_list.append('\t\t</s:Table>')
    chronect_list.append('\t</s:Worksheet>')
    chronect_list.append('\t<s:Worksheet s:Name="grdSettings">')
    chronect_list.append('\t\t<s:Table>')
    chronect_list.extend(2 * ['\t\t\t<s:Column s:Width="50" />'],)
    chronect_list.append(new_row)
    chronect_list.append(create_xml_cell('chkRepeatSchedule', 'String'))
    chronect_list.append(create_xml_cell(chk_repeat_schedule, 'String'))
    chronect_list.append(end_row)
    chronect_list.append(new_row)
    chronect_list.append(create_xml_cell('chkPrioritySchedule', 'String'))
    chronect_list.append(create_xml_cell(chk_priority_schedule, 'String'))
    chronect_list.append(end_row)
    chronect_list.append(new_row)
    chronect_list.append(create_xml_cell('chkOverlappedSchedule', 'String'))
    chronect_list.append(create_xml_cell(chk_overlapped_schedule, 'String'))
    chronect_list.append(end_row)
    chronect_list.append('\t\t</s:Table>')
    chronect_list.append('\t</s:Worksheet>')
    chronect_list.append('</s:Workbook>')

    chronect_input_written = False
    while not chronect_input_written:
        try:
            with open(chronect_input_file, 'w', newline = '', encoding = 'utf-16') as fout:
                i = 1
                for line in chronect_list:
                    if 's:Type="Number">_</s:Data>' in line:
                        line_list = line.split('s:Type="Number">_</s:Data>')
                        new_line = ''
                        for section in line_list[:-1]:
                            new_line = new_line + section + f's:Type="Number">{i}</s:Data>'
                            i += 1
                        new_line = new_line + line_list[-1]
                        line = ('').join(new_line)
                    fout.write(line + '\r\n')
        except PermissionError as error:
            filename = os.path.join(mydir, chronect_input_file)
            permission_error_popup(error, filename)
        else:
            chronect_input_written = True
        
    add_final_message(f'{chronect_input_file} Chronect inputfile written.')

# GUI functions
def combobox(question, dropdown_list, default_text = '', size = (5,4), tooltip = None):
    if params.SHOW_GUI:
        layout = [  [sg.Text(question.strip(), size = (30, None), tooltip = tooltip)],
                    [sg.Sizer(30), sg.Combo(dropdown_list, default_value = '_', key = 'result', size = size, enable_events = True)],
                    [sg.Button('OK', bind_return_key = True), sg.Text('', size = (23, None), right_click_menu= right)]
                    ]

        window = sg.Window('Choose parameters', layout, keep_on_top = True, return_keyboard_events = True, right_click_menu = right, use_default_focus = False)
        window.finalize()
        window.TKroot.focus_force()
        window.Element('result').SetFocus()
        window['result'].update(default_text)

        while True:
            try:
                event, values = window.read()

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

def add_final_message(message):
    global final_message
    
    print(message)
    if not message.strip():
        final_message.append(message.strip()) # Only strip if this leaves a non-zero length string
    else:
        final_message.append(message)

def final_popup(final_message, title):
    message = ('\n').join(final_message)
    
    if params.SHOW_GUI:
        button_column = [   [sg.Button('Open Chronect input', key = 'Quantos')],
                            [sg.Button('Close', bind_return_key = True)]
                            ]
        
        layout = [  [sg.Text(message), sg.Column(button_column)]
                    ]

        window = sg.Window(title, layout, keep_on_top = True, return_keyboard_events = True, right_click_menu = right)
        window.finalize()
        window.TKroot.focus_force()

        while True:
            try:
                event, _ = window.read()
                ''' You may need to add a delay or other function: loading the dialog by pressing Enter too slowly will trigger event == chr(13) '''
                if event in ['Close', sg.WIN_CLOSED, 'Escape:27', 'Exit', chr(13)]:
                    break
                elif event == 'Quantos':
                    print('Opening Chronect Quantos inputfile.')
                    open_chronect_file()
            except KeyboardInterrupt:
                break
        window.close()

# Open output files
def open_chronect_file():
    os.system(f'start excel "{os.path.join(mydir, chronect_input_file)}"')

# For faster debugging
def kill_excel():
    os.system('taskkill /f /im excel.exe')

sg.theme('Green')
right = ['right', ['Exit']]
final_message = []

pydir = params.pydir
mydir = pydir
os.chdir(mydir)

# Get rack type
''' All possible rack types should be in your rackparameters.
    Format: name(str), rows(int), cols(int)'''
rack_parameters = [['96', 8, 12], ['48', 6, 8], ['24 (4-mL)', 4, 6], ['24 (8-mL)', 4, 6]]
rack_type = get_rack_type()

plate_rows = list(string.ascii_uppercase)[:8] # Letters A-H
plate_cols = [str(x + 1) for x in range(24)]

for p in rack_parameters:
    if rack_type == p[0]:
        all_locations_no_zeroes = [(row + col) for row in plate_rows[:p[1]] for col in plate_cols[:p[2]]]
        break

# Get Chronect dosing tray location
chronect_dosing_tray = set_chronect_dosing_tray()

''' Here is the example input for the createSubstanceCSL function. You could automatically generate this from an eLN table, or user Excel input, for example '''
example_substances = [  {'Name': 't-BuBrettPhos', 'DoseLocations': ['', '1.0 mg', '', '1.0 mg', '', '1.0 mg', '', '1.0 mg', '', '1.0 mg', '', '1.0 mg', '', '1.0 mg', '', '1.0 mg', '', '1.0 mg', '', '1.0 mg', '', '', '', '', '', '', '', '']},
                        {'Name': 'CataCXium A', 'DoseLocations': ['2.1 mg', '2.1 mg', '2.1 mg', '2.1 mg', '2.1 mg', '2.1 mg', '2.1 mg', '2.1 mg', '0.5 mg', '', '0.5 mg', '', '0.5 mg', '', '0.5 mg', '', '', '', '', '', '', '', '', '']}
                        ]

csl_substances = []
for substance in example_substances:
    substance_csl = create_substance_csl(substance['Name'], substance['DoseLocations'])
    csl_substances.append(substance_csl)

# Write Chronect inputfile
if chronect_dosing_tray:
    chronect_input_file = 'inputfile.csl' ### Default name
    create_chronect_input()

if final_message:
    final_popup(final_message, 'Successful!')
