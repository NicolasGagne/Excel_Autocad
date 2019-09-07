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

        print('--------Event Start--------------')
        print(button, values, type(button), values['path'])
        print('--------Event End --------------')

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

            # pattern for number only
            pattern = "[^0-9.]"

            for i in range(df.shape[0]):

                df.iat[i, chainage_index] = str(df.iat[i, chainage_index])
                df.iat[i, chainage_index] = df.iat[i, chainage_index].strip()

                if window.find_element('comma'):
                    df.iat[i, chainage_index].replace(',', '.')

                try:
                    float(df.iat[i, chainage_index])

                except ValueError:
                    # Remove anything in front of the chainaige number
                    for x in range(len(df.iat[i, chainage_index])):

                        if not re.search(pattern, df.iat[i, chainage_index][x:]):
                            df.iat[i, chainage_index] = df.iat[i, chainage_index][x:]
                            break

            for i in range(df.shape[0]):

                df.iat[i, point_index] = str(df.iat[i, point_index])
                df.iat[i, point_index] = df.iat[i, point_index].strip()

                if window.find_element('comma'):
                    df.iat[i, point_index].replace(',', '.')

                try:
                    float(df.iat[i, point_index])


                except ValueError:
                    # Remove anything in front of the point number
                    for x in range(len(df.iat[i, point_index])):

                        if not re.search(pattern, df.iat[i, point_index][x:]):
                            df.iat[i, point_index] = df.iat[i, point_index][x:]
                            break

            # Select 2 column and Drop row with no Elevation
            selection = df.iloc[:, [chainage_index, point_index]]

            to_drop = []
            for i in range(selection.shape[0]):
                if selection.iat[i, -1] == 'nan' or selection.iat[i, 0] == 'nan':
                    to_drop.append(i)

            selection.drop(selection.index[to_drop], inplace=True)

            #Crée une 2e list en ordre pour comparaison
            sort_selection = selection.sort_values(selection.columns[0])
            clean = True
            for i in range(selection.shape[0]):
                if selection.iat[i, 0] == sort_selection.iat[i, 0]:
                    pass
                else:
                    # Message d'erreur avec l'idendtification du probleme
                    index_chainage_error = []
                    for x in range(df.shape[0]):
                        if df.iat[x, chainage_index] == selection.iat[i, 0] or df.iat[x, chainage_index] == sort_selection.iat[i, 0]:
                            index_chainage_error.append((x, df.iat[x, chainage_index]))

                    text = ""
                    for x in index_chainage_error:
                        text = text + 'Ordre de Chainage non valide; verifier fichier excel; ligne: ' +\
                               str(x[0] + row_index + 2) + '; Chainage:' + x[1]+ '\n'

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
