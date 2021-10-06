import datetime
import sys
from pprint import pprint

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor

sys.path.insert(1, 'C:\\Users\\timofey.inozemtsev\\PycharmProjects\\dll_power')

from dll_power import CANMarathon
from PyQt5 import QtWidgets, QtGui
from PyQt5.QtWidgets import QTableWidgetItem, QComboBox

import CANAnalyzer_ui
import pandas as pandas


def get_address(name: str):
    for param in params_list:
        if str(name) == str(param['name']):
            return int(param['address'])
    return 'nan'


def check_param(address: int, value):  # если новое значение - часть списка, то
    int_type_list = ['UINT32', 'UINT16', 'INT32', 'INT16', 'DATE']
    for param in params_list:
        if str(param['address']) != 'nan':
            if param['address'] == address:  # нахожу нужный параметр
                if str(param['editable']) != 'nan':  # он должен быть изменяемым
                    if param['type'] in int_type_list:  # если он в списке интов
                        if value.isdigit:  # и переменная - число
                            value = int(value)
                            if int(param['max']) >= value >= int(param['min']):  # причём это число в зоне допустимого
                                return value  # ну тогда так у ж и быть - отдаём это число
                            else:
                                print(f"param {value} is not in range from {param['min']} to {param['max']}")
                        else:
                            print(f"wrong type of param {type(value)}")
                    else:

                        # отработка попадания значения из списка STR и UNION
                        print(f"wrong is not numeric {param['type']}")
                else:
                    print(f"can't change param {param['name']}")

    return 'nan'


def set_param(address: int, value: int):
    global wr_err
    address = int(address)
    value = int(value)
    data = [value & 0xFF,
            ((value & 0xFF00) >> 8),
            0, 0,
            address & 0xFF,
            ((address & 0xFF00) >> 8),
            0x2B, 0x10]
    print(' Trying to set param in address ' + str(address) + ' to new value ' + str(value))
    if marathon.can_write(0x4F5, data):
        print(f'Successfully updated param in address {address} into devise')
        for param in params_list:
            if param['address'] == address:
                param['value'] = value
                print(f'Successfully written new value {value} in {param["name"]}')
                return True
    else:
        wr_err = "can't write param into device"
        return False


def get_param(address):
    global wr_err
    request_iteration = 3
    address = int(address)
    LSB = address & 0xFF
    MSB = ((address & 0xFF00) >> 8)
    for i in range(request_iteration):  # на случай если не удалось с первого раза поймать параметр,
        # делаем ещё request_iteration запросов
        data = marathon.can_request(0x4F5, 0x4F7, [0, 0, 0, 0, LSB, MSB, 0x2B, 0x03])
        if data:
            return (data[1] << 8) + data[0]
    wr_err = "can't read answer"
    return 'None'


def get_all_params():
    for param in params_list:
        if str(param['address']) != 'nan':
            param['value'] = get_param(address=int(param['address']))
    if not wr_err:
        return True
    print(wr_err)
    return False


def save_all_params():
    if get_all_params():
        file_name = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        file_name = 'Burr-30_' + file_name + '.xlsx'
        pandas.DataFrame(params_list).to_excel(file_name, index=False)
        print(' Save file success')
        return True
    print('Fail save file')
    return False


class ExampleApp(QtWidgets.QMainWindow, CANAnalyzer_ui.Ui_MainWindow):
    name_col = 0
    desc_col = 1
    value_col = 2
    combo_col = 3
    unit_col = 4

    def __init__(self):
        super().__init__()
        self.setupUi(self)  # Это нужно для инициализации нашего дизайна

    def list_of_params(self, item):
        global wr_err
        self.params_table.itemChanged.disconnect()
        self.params_table.setRowCount(0)
        self.params_table.setRowCount(len(bookmark_dict[item.text()]))
        row = 0

        for par in bookmark_dict[item.text()]:

            name_Item = QTableWidgetItem(par['name'])
            name_Item.setFlags(name_Item.flags() & ~Qt.ItemIsEditable)
            self.params_table.setItem(row, self.name_col, name_Item)
            if str(par['description']) != 'nan':
                description = str(par['description'])
            else:
                description = ''
            description_Item = QTableWidgetItem(description)
            self.params_table.setItem(row, self.desc_col, description_Item)

            if str(par['unit']) != 'nan':
                unit = str(par['unit'])
            else:
                unit = ''
            unit_Item = QTableWidgetItem(unit)
            unit_Item.setFlags(unit_Item.flags() & ~Qt.ItemIsEditable)
            self.params_table.setItem(row, self.unit_col, unit_Item)

            value = get_param(int(par['address']))

            if wr_err:
                wr_err = ''
            else:
                if str(par['strings']) == 'nan':
                    value_Item = QTableWidgetItem(str(value))
                    if str(par['editable']) != 'nan':
                        value_Item.setFlags(value_Item.flags() | Qt.ItemIsEditable)
                        value_Item.setBackground(QColor('#D7FBFF'))
                    else:
                        value_Item.setFlags(value_Item.flags() & ~Qt.ItemIsEditable)
                    self.params_table.setItem(row, self.value_col, value_Item)
                else:
                    string_dict = {}
                    for item in par['strings'].strip().split(';'):
                        if item:
                            it = item.split('-')
                            string_dict[int(it[0].strip())] = it[1]
                    combo_list = QComboBox()
                    for st in string_dict.values():
                        combo_list.addItem(st)
                    if value != 'None':
                        combo_list.setCurrentIndex(value)
                    if str(par['editable']) != 'nan':
                        combo_list.setEditable(True)
                        pal = combo_list.palette()
                        from PyQt5.QtGui import QPalette
                        pal.setColor(QPalette.setColor(), QColor('#D7FBFF'))
                        combo_list.setPalette(pal)
                    else:
                        combo_list.setEditable(False)
                    self.params_table.setCellWidget(row, self.value_col, combo_list)

            row += 1
        self.params_table.resizeColumnsToContents()
        self.params_table.itemChanged.connect(window.save_item)

    def save_item(self, item):
        new_value = item.text()
        print(new_value)
        name_param = self.params_table.item(item.row(), self.name_col).text()
        if item.column() == self.value_col:
            address_param = get_address(name_param)
            if str(address_param) != 'nan':
                value = check_param(address_param, new_value)
                if str(value) != 'nan':  # прошёл проверку
                    if set_param(address_param, value):

                        return True
                    else:
                        print("Can't write param into device")
                else:
                    print("Param isn't in available range")
            else:
                print("Can't find this value")
            return False
        elif item.column() == self.desc_col:
            for param in params_list:
                if name_param == str(param['name']):
                    param['description'] = new_value
                    return True
            print("Can't find this value")
            return False
        else:
            print("It's impossible!!!")
            return False


app = QtWidgets.QApplication([])
window = ExampleApp()  # Создаём объект класса ExampleApp

excel_data_df = pandas.read_excel('burr_30_forw_v31_27072021.xls')
params_list = excel_data_df.to_dict(orient='records')
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
# window.installEventFilter(window.params_table)
window.params_table.itemChanged.connect(window.save_item)
window.pushButton.clicked.connect(save_all_params)
window.list_bookmark.itemClicked.connect(window.list_of_params)
window.params_table.resizeColumnsToContents()

window.show()  # Показываем окно
app.exec_()  # и запускаем приложение
