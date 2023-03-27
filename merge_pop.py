import PySimpleGUI as sg
import pandas as pd
from typing import Optional
import constants

def merge(data_file:str, data_columns: list[str], data_sheet: str,
          translation_file: str, translation_columns: list[str], translation_sheet: str,
        ) -> bool:
    data_column = None
    translation_column = None
    output = None
    
    layout = [
        [
          sg.Text('Выберите столбец для соединения: '),
        ],
        [
            sg.Listbox(data_columns,
                       size=(30,10),
                       enable_events=True,
                       key=constants.CHOOSE_DATA_COLUMN,
                       ),
            sg.VerticalSeparator(),
            sg.Listbox(translation_columns,
                       size=(30,10),
                       enable_events=True,
                       key=constants.CHOOSE_TRANSLATION_COLUMN),
        ],
        [
            sg.In(disabled = True, text_color='black', enable_events=True, key=constants.OUTPUT_FILE),
            sg.FileSaveAs(key=constants.SAVE_FILE_AS,  initial_folder='.', file_types=(("Excel File", "*.xlsx"),)),
        ],
        [
            sg.Button(constants.SAVE_BUTTON),
            sg.Button(constants.CANCEL_BUTTON),
        ],
    ]
    
    window = sg.Window('Choose columns', layout)
    
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == constants.CANCEL_BUTTON:
            break
        
        if event == constants.CHOOSE_DATA_COLUMN:
            data_column = values[constants.CHOOSE_DATA_COLUMN][0]
        elif event == constants.CHOOSE_TRANSLATION_COLUMN:
            translation_column = values[constants.CHOOSE_TRANSLATION_COLUMN][0]
        elif event == constants.OUTPUT_FILE:
            output = values[constants.OUTPUT_FILE]
        elif event == constants.SAVE_BUTTON:
            if data_column and translation_column and output:
                data_df = pd.read_excel(data_file, sheet_name=data_sheet)
                data_df = data_df[data_columns]
                
                translation_df = pd.read_excel(translation_file, sheet_name=translation_sheet)
                translation_df = translation_df[translation_columns]
                translation_df = translation_df.rename(columns = {translation_column: data_column})
                
                data_df = data_df.merge(translation_df, on=data_column, how='left')
                
                data_df.to_excel(output, index=False)
                
                sg.popup('Сохранен')
                
                break
            else:
                sg.popup('Столбцы или файл не выбраны')
                
                
    
    window.close()
