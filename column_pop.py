import PySimpleGUI as sg
import pandas as pd
from typing import Optional
import constants

def getExcelColumns(filename: str) -> tuple[str, Optional[list[str]]]:
    sheet_name = ''
    column_names = []
    xl = pd.ExcelFile(filename)
    sheet_names = list(xl.sheet_names) 
    
    layout = [
        [
          sg.Text('Выберите Sheet: '),
          sg.Listbox(sheet_names, size=(10,4), enable_events=True, key=constants.CHOOSE_BUTTON),
        ],
        [sg.Column([[]], key=constants.DATA_COLUMNS)],
        [sg.Button(constants.OK_BUTTON)],
    ]
    
    window = sg.Window('Choose columns', layout)
    
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            break
        
        if event == constants.CHOOSE_BUTTON:
            sheet_name = values[constants.CHOOSE_BUTTON][0]
            try:
                df = pd.read_excel(filename, sheet_name=sheet_name, nrows=1)
                column_names = [col for col in list(df.columns)]
                
                layout = [
                    [
                        sg.Text('Выберите Sheet: '),
                        sg.Listbox(sheet_names, size=(10,4), enable_events=True, key=constants.CHOOSE_BUTTON),
                    ],
                    [sg.Column([
                            [sg.Checkbox(col, default=True, key=constants.CHOOSE_BUTTON)] for col in column_names
                        ])],
                    [sg.Button(constants.OK_BUTTON)],
                ]
                
                
                window.close()
                window = sg.Window('Choose columns', layout)
                
            except Exception as e:
                sg.popup('Ошибка при открытии')
        elif event == constants.OK_BUTTON:
            result = []
            
            for i in range(len(column_names)):
                try:
                    if values[constants.CHOOSE_BUTTON + str(i)]:
                        result.append(
                            column_names[i]
                        )
                except:
                    break
            
            window.close()
            
            return sheet_name, result
        
        print('You entered ', values[constants.CHOOSE_BUTTON][0])
