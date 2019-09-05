"""
File To host the GUI interface with the User
"""

import PySimpleGUI as sg
import pandas as pd
from xlrd.biffh import XLRDError
import re
from string import ascii_uppercase
from Layout import layout

def main():

    window = sg.Window('Excel a Liste de point ').Layout(layout)

    while True:
        button, values = window.Read()

        print(button, values, type(button), values['path'])

        if button == 'submit_file':
            try:

                df = pd.read_excel(values['path'])

                x = df.to_string(header=True,
                                 index=False,
                                 index_names=False).split('\n')
                rows_list = [','.join(ele.split()) for ele in x]

                # Add Row number
                for i in range(len(rows_list)):
                    rows_list[i] = str(i + 1) + '-  ' + rows_list[i]

                if rows_list[0] == 'Empty,DataFrame':
                    sg.PopupError('Fichier vide')
                    button = 'reset'

                else:
                    window.find_element('row').Update(disabled=False)
                    window.find_element('row').Update(values=rows_list)
                    window.find_element('submit_row').Update(disabled=False)



            except FileNotFoundError:
                sg.PopupError('Fichier introuvable')
                button = 'reset'

            except XLRDError:
                sg.PopupError('Format non supporté')
                button = 'reset'

        if button == 'submit_row':

            row_list_original = values['row'][0][4:].split(",")

            # Modified list to avoid duplicate
            row_list = []

            for (letter, elem) in zip(ascii_uppercase, row_list_original):
                row_list.append(letter + '- ' + elem)

            # Activate Button for Next Step
            window.find_element('row_box').Update(disabled=False)
            window.find_element('chainage_box').Update(disabled=False)
            window.find_element('row_box').Update(values=row_list)
            window.find_element('chainage_box').Update(values=row_list)

        if values['chainage_box'] != values['row_box'] and values['chainage_box'][0][-3:] != 'NaN' and values['row_box'][0][-3:] != 'NaN':
            window.find_element('save_as').Update(disabled=False)
            window.find_element('comma').Update(disabled=False)

        else:
            window.find_element('save_as').Update(disabled=True)
            window.find_element('comma').Update(disabled=True)

        if button == 'save_as':

            row_index = rows_list.index(values['row'][0])

            # drop all row above the title column
            for i in range(0,row_index, 1):
                df.drop(i, inplace=True)

            chainage_index = row_list.index(values['chainage_box'][0])
            point_index = row_list.index(values['row_box'][0])

            selection = df.iloc[:, [chainage_index, point_index]]
            # Patten for numbers only
            pattern = "[^0-9.]"

            # Drop row with no Elevation
            selection.dropna(inplace=True)

            for i in range(selection.shape[0]):
                for j in range(selection.shape[1]):
                    selection.iat[i, j] = str(selection.iat[i, j])
                    selection.iat[i, j] = selection.iat[i, j].strip()

                    if window.find_element('comma'):
                        selection.iat[i, j].replace(',', '.')

                    try:
                        float(selection.iat[i, j])

                    except ValueError:
                        # Remove anything in front of the chainaige number
                        for x in range(len(selection.iat[i, j])):

                            if not re.search(pattern, selection.iat[i, j][x:]):
                                selection.iat[i, j] = selection.iat[i, j][x:]
                                break

            #Crée une 2e list en ordre pour comparaison
            sort_selection = selection.sort_values(selection.columns[0])
            clean = True
            for i in range(selection.shape[0]):

                if selection.iat[i, 0] == sort_selection.iat[i, 0]:
                    pass
                else:
                    # Message d'erreur avec l'idendtification du probleme
                    text = 'Ordre de Chainage non valide; verifier fichier excel; ligne: ' + \
                           str(i + row_index + 1) + '; Donnee:' + selection.iat[i, 0]
                    sg.PopupError(text)
                    clean = False
                    button = 'reset'
                    break

            if clean:
                save_path = values['save_as']

                # Vérifie l'extention du ficher pour cfm .txt
                if save_path[-4:] != '.txt':
                    sg.PopupOK('Extemtion de ficher invalide, sera changer pour ".txt"')
                    index = save_path.find('.')
                    if index < 0:
                        index = len(save_path)

                    save_path = save_path[:index] + '.txt'

                #Sauvegard le fichier
                selection.to_csv(save_path, header=False, index=False, sep=' ', mode='w')

                # Remove empty line form the end of the file

                with open(save_path, 'r') as f:
                    data = f.read()
                    with open(save_path, 'w') as w:
                        w.write(data[:-1])

                sg.PopupOK('Fichier Sauvegarder')
                button = 'reset'

        if button == 'reset':
            # Réinitalise la fenetre
            window.find_element('row').Update(values=[])
            window.find_element('row').Update(disabled=True)
            window.find_element('submit_row').Update(disabled=True)
            window.find_element('row_box').Update(values=[])
            window.find_element('row_box').Update(disabled=True)
            window.find_element('chainage_box').Update(values=[])
            window.find_element('chainage_box').Update(disabled=True)
            window.find_element('save_as').Update(disabled=True)
            window.find_element('comma').Update(disabled=True)

        if button == 'Cancel' or button == 'Close':
            # Ferme l'application
            break
