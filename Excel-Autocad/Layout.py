"""
Layout file hold the layout as the window
"""

import PySimpleGUI as sg

layout = [
    [sg.Text('Ficher a convertir'), sg.Input(key='path'), sg.FileBrowse(key="find_path")],
    [sg.Submit(key='submit_file', size=[12, 1])],
    [sg.Text('_' * 95)],
    [sg.Text('Configuration', size=(15, 1))],
    [sg.Text('Title Row', size=(10, 1)), sg.Listbox(key='row', values=(''), size=(75, 7), disabled=True)],
    [sg.Submit(key='submit_row', disabled=True, size=[12, 1])],
    [sg.Text('_' * 95)],
    [sg.Text('Chainage', size=(10, 1)), sg.Listbox(key='chainage_box', values=(''), size=(30, 7), disabled=True, enable_events=True),
     sg.Text('Point', size=(7, 1)), sg.Listbox(key='row_box', values=(''), size=(30, 7), disabled=True, enable_events=True)],
    [sg.Text('Option:', size=(15, 1))],
    [sg.Checkbox('Check . vs ,', default=True, disabled=True, key='comma')],
    [sg.FileSaveAs('Convertir', disabled=True, key='save_as', enable_events=True, size=[12, 1])],
    [sg.Text('_' * 95)],
    [sg.Submit(key='reset', button_text='Reset', size=[12, 1]), sg.CloseButton(button_text='Close', size=[12, 1]),
     sg.Cancel(size=[12, 1])],
]