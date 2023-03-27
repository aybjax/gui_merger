import PySimpleGUI as sg
import pandas as pd
import constants
import column_pop
import merge_pop

sg.theme('DarkAmber')   # Add a touch of color

# All the stuff inside your window.

layout = [
            [sg.Text('Выбери excel файлы: ')],
            
            # [sg.HorizontalSeparator()],
            
            [sg.Text('     с данными'),
                sg.In(disabled = True, text_color='black', enable_events=True, key=constants.DATA_FILE),
                sg.FileBrowse(initial_folder='.', file_types=(("Excel File", "*.xlsx"),))],
            
            [sg.HorizontalSeparator()],
            
            [sg.Text('с переводами'),
                sg.In(disabled = True, text_color='black', enable_events=True, key=constants.TRANSLATOR_FILE),
                sg.FileBrowse(initial_folder='.', file_types=(("Excel File", "*.xlsx"),))],
            
            [sg.Button(constants.OK_BUTTON), sg.Button(constants.CANCEL_BUTTON)],
        ]

# Create the Window
window = sg.Window('Excel left joiner', layout)
data_file: str = None
translation_file: str = None

data_sheet: str = None
translation_sheet: str = None

data_columns: list[str] = None
translation_columns: list[str] = None
# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == constants.CANCEL_BUTTON: # if user closes window or clicks cancel
        break
    
    if event == constants.DATA_FILE:
        data_sheet, data_columns = column_pop.getExcelColumns(values[constants.DATA_FILE])
        data_file = values[constants.DATA_FILE]
    if event == constants.TRANSLATOR_FILE:
        translation_sheet, translation_columns = column_pop.getExcelColumns(values[constants.TRANSLATOR_FILE])
        translation_file = values[constants.TRANSLATOR_FILE]
    if event == constants.OK_BUTTON:
        print(f'data_file = {data_file}; data_columns = {data_columns}; translation_columns = {translation_columns}; translation_file = {translation_file}; ')
        if data_file and data_columns and translation_file and translation_columns and data_sheet and translation_sheet:
            window.close()
            
            merge_pop.merge(
                data_file=data_file, data_columns=data_columns, data_sheet=data_sheet,
                translation_file=translation_file, translation_columns=translation_columns, translation_sheet=translation_sheet,
            )
        else:
            sg.popup('Файлы или sheets не выбраны')

window.close()
