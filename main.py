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

Front_Wheel = 0x4F5
Rear_Wheel = 0x4F6
current_wheel = Front_Wheel

often_used_params = {
    'zone_of_insensitivity': {'scale': 100,
                              'value': 0,
                              'address': 103,
                              'min': 1,
                              'max': 5},
    'warning_temperature': {'scale': 1,
                            'value': 0,
                            'address': 104,
                            'min': 30,
                            'max': 80},
    'warning_current': {'scale': 10,
                        'value': 0,
                        'address': 105,
                        'min': 10,
                        'max': 60},
    'cut_off_current': {'scale': 10,
                        'value': 0,
                        'address': 403,
                        'min': 20,
                        'max': 80},
    'current_wheel': {'scale': 'nan',
                      'value': 0,
                      'address': 35},
    'byte_order': {'scale': 'nan',
                   'value': 0,
                   'address': 109},
}


def update_often_used():
    pass


def update_param():
    if window.tab_burr.currentWidget() == window.often_used_params:
        window.best_params()
    elif window.tab_burr.currentWidget() == window.all_params:
        param_list_clear()
        window.list_of_params(window.list_bookmark.currentItem())


def param_list_clear():
    for param in params_list:
        param['value'] = 'nan'
    return True


def rb_clicked():
    global current_wheel
    if window.radioButton.isChecked():
        current_wheel = Front_Wheel
    elif window.radioButton_2.isChecked():
        current_wheel = Rear_Wheel
    update_param()
    return True


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
                    if value.isdigit:  # и переменная - число
                        value = int(value)
                        # if int(param['max']) >= value >= int(param['min']):  # причём это число в зоне допустимого
                        return value  # ну тогда так у ж и быть - отдаём это число
                        # else:
                        #     print(f"param {value} is not in range from {param['min']} to {param['max']}")
                    else:
                        # отработка попадания значения из списка STR и UNION
                        print(f"value is not numeric {param['type']}")
                        string_dict = {}
                        for item in param['strings'].strip().split(';'):
                            if item:
                                it = item.split('-')
                                string_dict[it[1].strip()] = int(it[0].strip())
                        return string_dict[value.strip()]
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
    if marathon.can_write(current_wheel, data):
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
        data = marathon.can_request(current_wheel, current_wheel + 2, [0, 0, 0, 0, LSB, MSB, 0x2B, 0x03])
        if data:
            return (data[1] << 8) + data[0]
    wr_err = "can't read answer"
    return 'nan'


def get_all_params():
    for param in params_list:
        if str(param['address']) != 'nan':
            if str(param['value']) == 'nan':
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
    combo_col = 999
    unit_col = 3

    def __init__(self):
        super().__init__()
        self.setupUi(self)  # Это нужно для инициализации нашего дизайна

    def best_params(self):
        for name, par in often_used_params.items():
            par['value'] = get_param(int(par['address']))
            if par['scale'] != 'nan':
                slider = getattr(self, name)
                slider.setMinimum(par['min'])
                slider.setMaximum(par['max'])
                if par['value'] != 'nan':
                    param = par['value'] / par['scale']
                else:
                    param = par['max']
                print(f'Param {name} is {param}')
                slider.setValue(param)
                # self.warning_current.setValue(param)

    def list_of_params(self, item):
        item = bookmark_dict[item.text()]
        global wr_err
        self.params_table.itemChanged.disconnect()
        self.params_table.setRowCount(0)
        self.params_table.setRowCount(len(item))
        row = 0

        for par in item:

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

            if str(par['value']) == 'nan':
                value = get_param(int(par['address']))
                par['value'] = value
            else:
                value = par['value']

            if wr_err:
                wr_err = ''
            else:
                # combo_list = QComboBox()
                # if str(par['strings']) != 'nan':  # если у нас есть список
                #     string_dict = {}
                #     for item in par['strings'].strip().split(';'):
                #         if item:
                #             it = item.split('-')
                #             string_dict[int(it[0].strip())] = it[1]
                #     for st in string_dict.values():
                #         combo_list.addItem(st)
                #     if value != 'None':
                #         combo_list.setCurrentIndex(value)

                value_Item = QTableWidgetItem(str(value))

                if str(par['editable']) != 'nan':
                    value_Item.setFlags(value_Item.flags() | Qt.ItemIsEditable)
                    value_Item.setBackground(QColor('#D7FBFF'))
                else:
                    value_Item.setFlags(value_Item.flags() & ~Qt.ItemIsEditable)

                if str(par['strings']) != 'nan':
                    value_Item.setStatusTip(str(par['strings']))
                    value_Item.setToolTip(str(par['strings']))

                self.params_table.setItem(row, self.value_col, value_Item)

            row += 1
        self.params_table.resizeColumnsToContents()
        self.params_table.itemChanged.connect(window.save_item)

    def save_item(self, item):
        new_value = item.text()
        if not new_value:
            print('Value is empty')
            return False
        name_param = self.params_table.item(item.row(), self.name_col).text()
        if item.column() == self.value_col:
            address_param = get_address(name_param)
            if str(address_param) != 'nan':
                value = check_param(address_param, new_value)
                if str(value) != 'nan':  # прошёл проверку
                    for i in range(3):
                        if set_param(address_param, value):
                            check_value = get_param(address_param)
                            if check_value == value:
                                print('Checked changed value - OK')
                                self.params_table.item(item.row(), self.value_col).setBackground(QColor('green'))
                                return True
                            else:
                                self.params_table.item(item.row(), self.value_col).setBackground(QColor('red'))
                        else:
                            print("Can't write param into device")
                    if address_param == 35:  # если произошла смена рейки, нужно поменять адреса
                        if self.radioButton.isChecked():
                            self.radioButton_2.setChecked()
                        else:
                            self.radioButton.setChecked()
                        rb_clicked()
                        return True
                    return False
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
window.radioButton.toggled.connect(rb_clicked)
window.radioButton_2.toggled.connect(rb_clicked)

window.params_table.itemChanged.connect(window.save_item)
window.pushButton.clicked.connect(save_all_params)
window.pushButton_2.clicked.connect(update_param)

window.list_bookmark.itemClicked.connect(window.list_of_params)
window.params_table.resizeColumnsToContents()

window.show()  # Показываем окно
app.exec_()  # и запускаем приложение
