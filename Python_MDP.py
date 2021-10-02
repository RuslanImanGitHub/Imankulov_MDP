# Загрузка библиотек
import pandas as pd
import sys
import win32com.client
import PySimpleGUI as sg
import Singleton

rastr = win32com.client.Dispatch('Astra.Rastr')

if __name__ == '__main__':
    # Графический интерфейс
    layout = [
        [sg.Text('Шаблон режима'), sg.InputText(), sg.FileBrowse()
         ],
        [sg.Text('Файл режима'), sg.InputText(), sg.FileBrowse(),
         ],
        [sg.Text('Шаблон сечения'), sg.InputText(), sg.FileBrowse()
         ],
        [sg.Text('Файл сечения'), sg.InputText(), sg.FileBrowse(),
         ],
        [sg.Text('Файл аварий'), sg.InputText(), sg.FileBrowse()
         ],
        [sg.Text('Файл траектории'), sg.InputText(), sg.FileBrowse(),
         ],
        [sg.Text('Шаблон траектории'), sg.InputText(), sg.FileBrowse(),
         ],
        [sg.Submit(), sg.Cancel()]
    ]
    window = sg.Window('Расчет МДП', layout)

    def check_file_path(file):
        """
        выполняет провверку наличия пути к файлу
        :param file: путь к файлу
        """
        if file is None:
            print('Укажите путь ко всем файлам')
            allow = 0

    while True:
        event, values = window.read()
        if event in (None, 'Exit', 'Cancel'):
            break
        if event == 'Submit':
            reg_shab = reg = sech_shab = sech =\
                fault = traject = traject_shab = allow = None

            reg_shab = values[0]
            reg = values[1]
            sech_shab = values[2]
            sech = values[3]
            fault = values[4]
            traject = values[5]
            traject_shab = values[6]
            allow = 1
            check_file_path(reg_shab)
            check_file_path(reg)
            check_file_path(sech_shab)
            check_file_path(sech)
            check_file_path(fault)
            check_file_path(traject)
            check_file_path(traject_shab)
            if allow == 1:
                Singleton.trajectory_loading(traject, traject_shab)
                Singleton.flowgate_loading(sech, sech_shab)
                faults = Singleton.faults_loading(fault)

                result_data = Singleton.criteria1(
                    reg, reg_shab, 0, traject_shab, sech_shab)
                result_data = Singleton.criteria2(
                    reg, reg_shab, 0, result_data, traject_shab, sech_shab)
                result_data = Singleton.criteria3(
                    reg, reg_shab, 0, result_data, traject_shab, sech_shab,
                    faults)
                result_data = Singleton.criteria4(
                    reg, reg_shab, 0, result_data, traject_shab, sech_shab,
                    faults)
                result_data = Singleton.criteria5(
                    reg, reg_shab, 0, result_data, traject_shab, sech_shab)
                result_data = Singleton.criteria6(
                    reg, reg_shab, 0, result_data, traject_shab, sech_shab,
                    faults)
    window.close()

