import datetime
import sys
from pprint import pprint

from PyQt5.QtCore import Qt

sys.path.insert(1, 'C:\\Users\\timofey.inozemtsev\\PycharmProjects\\dll_power')

from dll_power import CANMarathon
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QTableWidgetItem

import CANAnalyzer_ui
import pandas as pandas


def get_param(address):
    global wr_err
    address = int(address)
    LSB = address & 0xFF
    MSB = ((address & 0xFF00) >> 8)
    # print(hex(LSB) + " " + hex(MSB))
    # if not marathon.can_write(0x4F5, [0, 0, 0, 0, LSB, MSB, 0x2B, 0x03]):
    #     wr_err = "can't send request"
    #     return
    data = marathon.can_request(0x4F5, 0x4F7, [0, 0, 0, 0, LSB, MSB, 0x2B, 0x03])
    # print((data[1] << 8) + data[0])
    if data:
        return (data[1] << 8) + data[0]
    else:
        wr_err = "can't read answer"
        return 'None'


class ExampleApp(QtWidgets.QMainWindow, CANAnalyzer_ui.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)  # Это нужно для инициализации нашего дизайна

    def list_of_params(self, item):
        global wr_err
        self.params_table.setRowCount(0)
        self.params_table.setRowCount(len(bookmark_dict[item.text()]))
        row = 0

        for par in bookmark_dict[item.text()]:
            name_Item = QTableWidgetItem(par['name'])
            name_Item.setFlags(name_Item.flags() & ~Qt.ItemIsEditable)
            self.params_table.setItem(row, 0, name_Item)
            if str(par['description']) != 'nan':
                description = str(par['description'])
            else:
                description = ''
            description_Item = QTableWidgetItem(description)
            description_Item.setFlags(description_Item.flags() & ~Qt.ItemIsEditable)
            self.params_table.setItem(row, 1, description_Item)
            print(str(par['unit']))
            if str(par['unit']) != 'nan':
                unit = str(par['unit'])
            else:
                unit = ''
            unit_Item = QTableWidgetItem(unit)
            unit_Item.setFlags(unit_Item.flags() & ~Qt.ItemIsEditable)
            self.params_table.setItem(row, 3, unit_Item)
            value = get_param(int(par['address']))
            print(value)
            value_Item = QTableWidgetItem(str(value))
            if str(par['editable']) != 'nan':
                if int(par['editable']):
                    value_Item.setFlags(value_Item.flags() | Qt.ItemIsEditable)
                else:
                    value_Item.setFlags(value_Item.flags() & ~Qt.ItemIsEditable)

            if wr_err:
                wr_err = ''
            else:
                self.params_table.setItem(row, 2, value_Item)

            row += 1
        self.params_table.resizeColumnsToContents()


app = QtWidgets.QApplication([])
window = ExampleApp()  # Создаём объект класса ExampleApp

excel_data_df = pandas.read_excel('burr_30_forw_v31_27072021.xls')
params_list = excel_data_df.to_dict(orient='records')
pprint(params_list)
bookmark_dict = {}
bookmark_list = []
prev_name = ''
wr_err = ''

for param in params_list:
    if param['code'].count('.') == 2:
        param['address'] = int(param['address'])
        bookmark_list.append(param)
    elif param['code'].count('.') == 1:
        bookmark_dict[prev_name] = bookmark_list
        bookmark_list = []
        if prev_name:
            window.list_bookmark.addItem(prev_name)
        prev_name = param['name']

marathon = CANMarathon()

window.list_bookmark.itemClicked.connect(window.list_of_params)
window.params_table.resizeColumnsToContents()

window.show()  # Показываем окно
app.exec_()  # и запускаем приложение


def set_param(address: int, value: int):
    global wr_err
    address = int(address)
    data = [value & 0xFF,
            ((value & 0xFF00) >> 8),
            0, 0,
            address & 0xFF,
            ((address & 0xFF00) >> 8),
            0x2B, 0x10]
    pprint(data)
    if marathon.can_write(0x4F5, data):
        for param in params_list:
            if param['address'] == address:
                param['value'] = value
                break
        return True
    else:
        wr_err = "can't write param into device"
        return False


def get_all_params():
    for param in params_list:
        if param['address'] != 'nan':
            param['value'] = get_param(address=int(param['address']))
    if not wr_err:
        return True
    return False


def save_all_params():
    if get_all_params():
        file_name = datetime.datetime.now().strftime("%Y-%m-%d_%H:%M:%S")
        file_name = 'Burr-30_' + file_name + '.xlsx'
        pandas.DataFrame(params_list).to_excel(file_name)
        return True
    return False
