# python包
from PyQt5.QtWidgets import (QDialog, QTableView, QPushButton, QVBoxLayout, QStyledItemDelegate,
                             QLineEdit, QWidget, QGridLayout, QHBoxLayout, QLabel, QComboBox,
                             QCheckBox, QMessageBox)
from PyQt5.QtGui import QStandardItemModel, QStandardItem, QRegExpValidator
from PyQt5.QtCore import Qt, QRegExp


# 自建包
from CodeTools.loggingTool import PotsLogger

# 类

class childDelegate(QStyledItemDelegate):
    def __init__(self):
        super().__init__()

    def createEditor(self, parent, option, index):
        editor = QLineEdit(parent)
        return editor

    def setEditorData(self, editor, index):
        model = index.model()
        text = model.data(index, Qt.DisplayRole)
        editor.setText(text)

    def setModelData(self, editor, model, index):
        text = editor.text()
        model.setData(index, text, Qt.DisplayRole)

    def updateEditorGeometry(self, editor, option, index):
        editor.setGeometry(option.rect)


class childQStandardItemModel(QStandardItemModel):

    def data(self, index, role=None):
        if role == Qt.TextAlignmentRole:
            return Qt.AlignCenter
        return QStandardItemModel.data(self, index, role)

class PotsChildWindow_manager():
    def __init__(self, mainWindow, dataVessel):
        self.logger = PotsLogger('PotsChildWindow_manager')
        self.logger.normalInfo('管理对象初始化')

        self.mainWindnow = mainWindow
        self.dataVessel = dataVessel

        self.logger.normalInfo('创建窗体存放列表及卡件信息字典')
        self.list_childWindow = []
        self.dict_carton_confirmed = {
            'name_carton': '',
            'nameMin': '',
            'suport': '',
            'numRoad': 0,
            # 以上均为字符串
            'wires': [],
            'thoroughfare_wire': []  # 内容为各通道接线，数量等同于通道数
        }

        self.logger.normalInfo('创建首选项窗体,并将窗体放入存放列表')
        obj_firstWindow = PotsChildWindow_first(self.mainWindnow, self, self.dataVessel, self.dict_carton_confirmed)
        self.list_childWindow.append(obj_firstWindow)
        self.logger.normalInfo('展示首选项窗体')
        obj_firstWindow.show()

    def formatFunc(self):
        self.logger.getFuncName('formatFunc')
        self.logger.normalInfo('创建下一线制窗体,并将窗体放入存放列表')
        obj_nextWindow = PotsChildWindow_wire(self.mainWindnow, self, self.dataVessel, self.dict_carton_confirmed)
        self.list_childWindow.append(obj_nextWindow)
        self.logger.varInfo('formatFunc->此时列表内容为:'+str(self.list_childWindow))
        self.logger.normalInfo('展示下一窗体')
        obj_nextWindow.show()

    def sureFunc(self):
        self.logger.getFuncName('sureFunc')
        self.logger.varInfo('sureFunc->此时列表内容为:' + str(self.list_childWindow))
        self.logger.normalInfo('将最终的数据写入configData')
        try:
            result = self.dataVessel.insertValueToCustomizeCartonSheet(self.dict_carton_confirmed)
            if result == 0:
                dlg = QMessageBox(self.mainWidnow)
                dlg.setIcon(QMessageBox.Critical)
                dlg.setText('该卡件已存在')
                dlg.show()
            else:
                ...
        except:
            self.logger.tranceInfo()

class PotsChildWindow_wire(QDialog):
    def __init__(self, mainWindow, manager, dataVessel, dict_carton_confirmed):
        super(PotsChildWindow_wire, self).__init__(parent=mainWindow)
        self.logger = PotsLogger(
            'PotsChildWindow_wire->' + str(dict_carton_confirmed['wires'][len(manager.list_childWindow) - 1]))
        self.logger.normalInfo('初始化')

        self.mainWidnow = mainWindow
        self.manager = manager
        self.dataVessel = dataVessel

        self.list_childWindow = manager.list_childWindow

        self.dict_carton_confirmed = dict_carton_confirmed

        self.logger.normalInfo('创建窗体')
        num_windows = len(self.list_childWindow)
        self.setWindowTitle('卡件自定义——'+self.dict_carton_confirmed['wires'][num_windows-1])
        """
        self.setWindowModality(Qt.ApplicationModal)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        """
        # 设置窗口大小
        self.setFixedSize(600, 250)

        self.logger.normalInfo('创建窗体控件')
        self.layout_v = QVBoxLayout()

        self.logger.normalInfo('表格')
        try:
            self.layout_h_wireView = QHBoxLayout()
            num_road = int(self.dict_carton_confirmed['numRoad'])
            self.model = childQStandardItemModel(num_road, 4, self)
            self.model.setHorizontalHeaderLabels(['+', '-', 'C', '补'])
            for row in range(0, num_road):
                for column in range(0, 4):
                    item = QStandardItem('')
                    # 设置每个位置的文本值
                    self.model.setItem(row, column, item)
            self.tableView = QTableView(self)
            self.tableView.setModel(self.model)
            delegate = childDelegate()
            self.tableView.setItemDelegate(delegate)
            self.layout_h_wireView.addWidget(self.tableView)

            self.layout_v.addLayout(self.layout_h_wireView)
        except:
            self.logger.tranceInfo()

        self.logger.normalInfo('按钮')
        self.layout_h_button = QHBoxLayout()
        self.button_formoat = QPushButton('下一步')
        self.button_sure = QPushButton('确定')
        self.button_consuor = QPushButton('取消')
        self.logger.normalInfo('根据情况判别按钮使能及绑定')
        self.button_sure.clicked.connect(self.sureFunc)
        self.button_sure.clicked.connect(self.close)
        self.button_formoat.clicked.connect(self.formatFunc)
        self.button_formoat.clicked.connect(self.close)
        self.logger.varInfo('使能按钮时，判别条件的状态->\n 线制数' + str(len(self.dict_carton_confirmed['wires'])) +
                            '窗口数' + str(len(self.list_childWindow)))
        if len(self.dict_carton_confirmed['wires']) == len(self.list_childWindow):
            self.button_formoat.setEnabled(False)
            self.button_sure.setEnabled(True)
        else:
            self.button_formoat.setEnabled(True)
            self.button_sure.setEnabled(False)
        self.button_consuor.clicked.connect(self.close)
        self.layout_h_button.addWidget(self.button_formoat)
        self.layout_h_button.addWidget(self.button_sure)
        self.layout_h_button.addWidget(self.button_consuor)

        self.layout_v.addLayout(self.layout_h_button)

        self.setLayout(self.layout_v)

    def formatFunc(self):
        self.logger.getFuncName('formatFunc')
        self.logger.normalInfo('formatFunc->'+'添加卡件数据到信息字典')
        self.logger.varInfo('formatFunc->添加前信息字典内容'+str(self.dict_carton_confirmed))
        num_road = int(self.dict_carton_confirmed['numRoad'])
        list_wireThoroughfare = []
        for row in range(0, num_road):
            list_oneThoroughfare = []
            for column in range(0, 4):
                index_C0 = self.model.index(row, column)
                data = self.model.data(index_C0, Qt.DisplayRole)
                if data == '' or data == ' ' or data is None:
                    continue
                else:
                    list_oneThoroughfare.append(data)
            list_wireThoroughfare.append(list_oneThoroughfare)
        self.dict_carton_confirmed['thoroughfare_wire'].append(list_wireThoroughfare)
        self.logger.varInfo('formatFunc->添加后信息字典内容' + str(self.dict_carton_confirmed))
        self.logger.normalInfo('进入下一窗体')
        self.manager.formatFunc()

    def sureFunc(self):
        self.logger.getFuncName('sureFunc')
        self.logger.getFuncName('formatFunc')
        self.logger.normalInfo('formatFunc->' + '添加卡件数据到信息字典')
        self.logger.varInfo('formatFunc->添加前信息字典内容' + str(self.dict_carton_confirmed))
        num_road = int(self.dict_carton_confirmed['numRoad'])
        list_wireThoroughfare = []
        for row in range(0, num_road):
            list_oneThoroughfare = []
            for column in range(0, 4):
                index_C0 = self.model.index(row, column)
                data = self.model.data(index_C0, Qt.DisplayRole)
                if data == '' or data == ' ' or data is None:
                    continue
                else:
                    list_oneThoroughfare.append(data)
            list_wireThoroughfare.append(list_oneThoroughfare)
        self.dict_carton_confirmed['thoroughfare_wire'].append(list_wireThoroughfare)
        self.logger.varInfo('formatFunc->添加后信息字典内容' + str(self.dict_carton_confirmed))
        self.logger.normalInfo('写入configData')
        self.manager.sureFunc()

class PotsChildWindow_first(QDialog):
    def __init__(self, mainWindow, manager, dataVessel, dict_carton_confirmed):
        super(PotsChildWindow_first, self).__init__(parent=mainWindow)
        self.logger = PotsLogger('PotsChildWindow_first->首选项窗体')
        self.logger.normalInfo('初始化')

        self.mainWidnow = mainWindow
        self.manager = manager
        self.dataVessel = dataVessel

        self.dict_carton_confirmed = dict_carton_confirmed

        self.logger.normalInfo('创建窗体')
        self.setWindowTitle('卡件自定义——首选项')
        """
        self.setWindowModality(Qt.ApplicationModal)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        """

        # 设置窗口大小
        self.setFixedSize(600, 250)

        self.logger.normalInfo('创建窗体控件')
        self.layout_v = QVBoxLayout()

        self.logger.normalInfo('卡件名称')
        self.layout_h_cartonName = QHBoxLayout()
        self.layout_h_cartonName_label = QHBoxLayout()
        self.layout_h_cartonName_editor = QHBoxLayout()
        self.lable_cartonName_str = QLabel('卡件名称')
        self.editor_cartonName_str = QLineEdit(self)
        self.layout_h_cartonName_label.addWidget(self.lable_cartonName_str)
        self.layout_h_cartonName_editor.addWidget(self.editor_cartonName_str)
        self.layout_h_cartonName.addLayout(self.layout_h_cartonName_label)
        self.layout_h_cartonName.addLayout(self.layout_h_cartonName_editor)

        self.logger.normalInfo('卡件类型')
        self.layout_h_nameMin = QHBoxLayout()
        self.layout_h_nameMin_label = QHBoxLayout()
        self.layout_h_nameMin_editor = QHBoxLayout()
        self.lable_nameMin_str = QLabel('卡件类型')
        self.editor_nameMin_str = QLineEdit(self)
        self.layout_h_nameMin_label.addWidget(self.lable_nameMin_str)
        self.layout_h_nameMin_editor.addWidget(self.editor_nameMin_str)
        self.layout_h_nameMin.addLayout(self.layout_h_nameMin_label)
        self.layout_h_nameMin.addLayout(self.layout_h_nameMin_editor)

        self.logger.normalInfo('是否支持高密度')
        self.layout_h_suport = QHBoxLayout()
        self.layout_h_suport_label = QHBoxLayout()
        self.layout_h_suport_editor = QHBoxLayout()
        self.lable_suport_str = QLabel('是否支持高密度')
        self.editor_suport_str = QComboBox(self)
        self.editor_suport_str.setEnabled(True)
        self.editor_suport_str.addItems(['True', 'False'])
        self.layout_h_suport_label.addWidget(self.lable_suport_str)
        self.layout_h_suport_editor.addWidget(self.editor_suport_str)
        self.layout_h_suport.addLayout(self.layout_h_suport_label)
        self.layout_h_suport.addLayout(self.layout_h_suport_editor)

        self.logger.normalInfo('通道数(偶数)')
        self.layout_h_numRoad = QHBoxLayout()
        self.layout_h_numRoad_label = QHBoxLayout()
        self.layout_h_numRoad_editor = QHBoxLayout()
        self.lable_numRoad_str = QLabel('通道数(偶数)')
        self.editor_numRoad_str = QLineEdit(self)
        reg = QRegExp('[a-zA-z0-9]+$')
        validator = QRegExpValidator(self)
        validator.setRegExp(reg)
        self.editor_numRoad_str.setValidator(validator)
        self.layout_h_numRoad_label.addWidget(self.lable_numRoad_str)
        self.layout_h_numRoad_editor.addWidget(self.editor_numRoad_str)
        self.layout_h_numRoad.addLayout(self.layout_h_numRoad_label)
        self.layout_h_numRoad.addLayout(self.layout_h_numRoad_editor)

        self.logger.normalInfo('复选框')
        self.layout_v_wire = QVBoxLayout()
        self.layout_h_wire = QHBoxLayout()
        self.lable_wire_str = QLabel('请选择所支持的线制')
        self.layout_h_wire_checks = QHBoxLayout()
        self.editor_wire_str1 = QCheckBox('2线制',self)
        self.editor_wire_str2 = QCheckBox('3线制', self)
        self.editor_wire_str3 = QCheckBox('4线制', self)
        self.editor_wire_str4 = QCheckBox('220V', self)
        self.editor_wire_str5 = QCheckBox('L', self)
        self.editor_wire_str6 = QCheckBox('H', self)
        self.editor_wire_str7 = QCheckBox('OC', self)
        self.editor_wire_str8 = QCheckBox('TTL', self)
        self.editor_wire_strNo = QCheckBox('取消', self)
        self.list_checks = []
        self.list_checks.append(self.editor_wire_str1)
        self.list_checks.append(self.editor_wire_str2)
        self.list_checks.append(self.editor_wire_str3)
        self.list_checks.append(self.editor_wire_str4)
        self.list_checks.append(self.editor_wire_str5)
        self.list_checks.append(self.editor_wire_str6)
        self.list_checks.append(self.editor_wire_str7)
        self.list_checks.append(self.editor_wire_str8)
        self.list_checks.append(self.editor_wire_strNo)
        self.editor_wire_str1.stateChanged.connect(self.checkedFunc)
        self.editor_wire_str2.stateChanged.connect(self.checkedFunc)
        self.editor_wire_str3.stateChanged.connect(self.checkedFunc)
        self.editor_wire_str4.stateChanged.connect(self.checkedFunc)
        self.editor_wire_str5.stateChanged.connect(self.checkedFunc)
        self.editor_wire_str6.stateChanged.connect(self.checkedFunc)
        self.editor_wire_str7.stateChanged.connect(self.checkedFunc)
        self.editor_wire_str8.stateChanged.connect(self.checkedFunc)
        self.editor_wire_strNo.stateChanged.connect(self.checkedFunc)
        self.layout_h_wire.addWidget(self.lable_wire_str)
        self.layout_h_wire_checks.addWidget(self.editor_wire_str1)
        self.layout_h_wire_checks.addWidget(self.editor_wire_str2)
        self.layout_h_wire_checks.addWidget(self.editor_wire_str3)
        self.layout_h_wire_checks.addWidget(self.editor_wire_str4)
        self.layout_h_wire_checks.addWidget(self.editor_wire_str5)
        self.layout_h_wire_checks.addWidget(self.editor_wire_str6)
        self.layout_h_wire_checks.addWidget(self.editor_wire_str7)
        self.layout_h_wire_checks.addWidget(self.editor_wire_str8)
        self.layout_h_wire_checks.addWidget(self.editor_wire_strNo)

        self.layout_v_wire.addLayout(self.layout_h_wire)
        self.layout_v_wire.addLayout(self.layout_h_wire_checks)

        self.logger.normalInfo('按钮')
        self.layout_h_button = QHBoxLayout()
        self.button_formoat = QPushButton('下一步')
        self.button_formoat.setEnabled(False)
        self.button_sure = QPushButton('确定')
        self.button_sure.setEnabled(False)
        self.button_consuor = QPushButton('取消')
        self.button_formoat.clicked.connect(self.formatFunc)
        self.button_formoat.clicked.connect(self.close)
        self.button_consuor.clicked.connect(self.close)
        self.layout_h_button.addWidget(self.button_formoat)
        self.layout_h_button.addWidget(self.button_sure)
        self.layout_h_button.addWidget(self.button_consuor)

        self.layout_v.addLayout(self.layout_h_cartonName)
        self.layout_v.addLayout(self.layout_h_nameMin)
        self.layout_v.addLayout(self.layout_h_suport)
        self.layout_v.addLayout(self.layout_h_numRoad)
        self.layout_v.addLayout(self.layout_v_wire)
        self.layout_v.addLayout(self.layout_h_button)

        self.setLayout(self.layout_v)

    def checkedFunc(self):
        self.logger.getFuncName('checkedFunc')
        self.logger.normalInfo('根据复选框按钮状态确定线制')
        self.logger.varInfo('复选改变线制前的制列表->'+str(self.dict_carton_confirmed['wires']))
        try:
            self.dict_carton_confirmed['wires'] = []
            for editor_check in self.list_checks:
                if editor_check.isChecked():
                    if self.list_checks.index(editor_check) == 0:
                        if editor_check.checkState() == Qt.Checked:
                            self.dict_carton_confirmed['wires'].append('2线制')
                        else:
                            if '2线制' in self.dict_carton_confirmed['wires']:
                                index = self.dict_carton_confirmed['wires'].index('2线制')
                                del self.dict_carton_confirmed['wires'][index]
                    elif self.list_checks.index(editor_check) == 1:
                        if editor_check.checkState() == Qt.Checked:
                            self.dict_carton_confirmed['wires'].append('3线制')
                        else:
                            if '3线制' in self.dict_carton_confirmed['wires']:
                                index = self.dict_carton_confirmed['wires'].index('3线制')
                                del self.dict_carton_confirmed['wires'][index]
                    elif self.list_checks.index(editor_check) == 2:
                        if editor_check.checkState() == Qt.Checked:
                            self.dict_carton_confirmed['wires'].append('4线制')
                        else:
                            if '4线制' in self.dict_carton_confirmed['wires']:
                                index = self.dict_carton_confirmed['wires'].index('4线制')
                                del self.dict_carton_confirmed['wires'][index]
                    elif self.list_checks.index(editor_check) == 3:
                        if editor_check.checkState() == Qt.Checked:
                            self.dict_carton_confirmed['wires'].append('220V')
                        else:
                            if '220V' in self.dict_carton_confirmed['wires']:
                                index = self.dict_carton_confirmed['wires'].index('220V')
                                del self.dict_carton_confirmed['wires'][index]
                    elif self.list_checks.index(editor_check) == 4:
                        if editor_check.checkState() == Qt.Checked:
                            self.dict_carton_confirmed['wires'].append('L')
                        else:
                            if 'L' in self.dict_carton_confirmed['wires']:
                                index = self.dict_carton_confirmed['wires'].index('L')
                                del self.dict_carton_confirmed['wires'][index]
                    elif self.list_checks.index(editor_check) == 5:
                        if editor_check.checkState() == Qt.Checked:
                            self.dict_carton_confirmed['wires'].append('H')
                        else:
                            if 'H' in self.dict_carton_confirmed['wires']:
                                index = self.dict_carton_confirmed['wires'].index('H')
                                del self.dict_carton_confirmed['wires'][index]
                    elif self.list_checks.index(editor_check) == 6:
                        if editor_check.checkState() == Qt.Checked:
                            self.dict_carton_confirmed['wires'].append('OC')
                        else:
                            if 'OC' in self.dict_carton_confirmed['wires']:
                                index = self.dict_carton_confirmed['wires'].index('OC')
                                del self.dict_carton_confirmed['wires'][index]
                    elif self.list_checks.index(editor_check) == 7:
                        if editor_check.checkState() == Qt.Checked:
                            self.dict_carton_confirmed['wires'].append('TTL')
                        else:
                            if 'TTL' in self.dict_carton_confirmed['wires']:
                                index = self.dict_carton_confirmed['wires'].index('TTL')
                                del self.dict_carton_confirmed['wires'][index]
                    elif self.list_checks.index(editor_check) == 8:
                        if editor_check.checkState() == Qt.Checked:
                            self.dict_carton_confirmed['wires'] = []
                            self.list_checks[0].setChecked(False)
                            self.list_checks[1].setChecked(False)
                            self.list_checks[2].setChecked(False)
                            self.list_checks[3].setChecked(False)
                            self.list_checks[4].setChecked(False)
                            self.list_checks[5].setChecked(False)
                            self.list_checks[6].setChecked(False)
                            self.list_checks[7].setChecked(False)
                            self.dict_carton_confirmed['wires'] = []
                        else:
                            pass
                else:
                    continue
            if self.dict_carton_confirmed['wires'] == []:
                self.button_formoat.setEnabled(False)
            else:
                self.button_formoat.setEnabled(True)
            self.logger.varInfo('复选改变线制后的制列表->' + str(self.dict_carton_confirmed['wires']))
        except:
            self.logger.tranceInfo()

    def formatFunc(self):
        self.logger.getFuncName('formatFunc')
        self.logger.normalInfo('赋值')
        try:
            if self.editor_cartonName_str.text() == None:
                self.dict_carton_confirmed['name_carton'] = ''
            else:
                self.dict_carton_confirmed['name_carton'] = self.editor_cartonName_str.text()

            if self.editor_nameMin_str.text() == None:
                self.dict_carton_confirmed['nameMin'] = ''
            else:
                self.dict_carton_confirmed['nameMin'] = self.editor_nameMin_str.text()

            if self.editor_suport_str.currentText() == None:
                self.dict_carton_confirmed['suport'] = ''
            else:
                self.dict_carton_confirmed['suport'] = self.editor_suport_str.currentText()

            if self.editor_numRoad_str.text() == None or self.editor_numRoad_str.text() == '':
                self.dict_carton_confirmed['numRoad'] = ''
            else:
                self.dict_carton_confirmed['numRoad'] = int(self.editor_numRoad_str.text())
        except:
            self.logger.tranceInfo()

        self.logger.normalInfo('判别是否应该进入下一页')
        if self.dict_carton_confirmed['name_carton'] == '' or \
                self.dict_carton_confirmed['nameMin'] == '' or self.dict_carton_confirmed['suport'] == '' or \
                self.dict_carton_confirmed['numRoad'] == '':
            dlg = QMessageBox(self.mainWidnow)
            dlg.setIcon(QMessageBox.Critical)
            dlg.setText('首选项界面所有选项必填')
            dlg.show()
        else:
            self.manager.formatFunc()



# 测试函数
def test():
    pass


# 开启动作
if __name__ == '__main__':
    test()
