import json
from pprint import pprint

from PyQt5 import QtWidgets
import CANAnalyzer_ui
import pandas as pandas


class ExampleApp(QtWidgets.QMainWindow, CANAnalyzer_ui.Ui_MainWindow):
    def __init__(self):
        # Это здесь нужно для доступа к переменным, методам
        # и т.д. в файле design.py
        super().__init__()
        self.setupUi(self)  # Это нужно для инициализации нашего дизайна


excel_data_df = pandas.read_excel('burr_30_forw_v31_27072021.xls')
params_list = excel_data_df.to_dict(orient='records')
bookmark_dict = {}
bookmark_list = []
#
# for param in params_list:
#     if param['code'].count('.') == 2:
#         param['address'] = int(param['address'])
#         bookmark_list.append(param)
#     elif param['code'].count('.') == 1:
#         bookmark_dict[param['name']] = bookmark_list
#         bookmark_list = []
# pprint(bookmark_dict)

app = QtWidgets.QApplication([])
# window = ExampleApp()  # Создаём объект класса ExampleApp
# window.show()  # Показываем окно
app.exec_()  # и запускаем приложение
