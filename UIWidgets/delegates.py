# python包
from PyQt5.QtWidgets import QStyledItemDelegate, QComboBox, QLineEdit, QMessageBox
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QStandardItem
import copy

# 自建包
from CodeTools.loggingTool import PotsLogger

# 类
class PotsStartFormViewDelegate(QStyledItemDelegate):
    def __init__(self, parent):
        super().__init__()
        self.logger = PotsLogger('PotsStartFormViewDelegate')
        self.logger.normalInfo('进行模件代理对象初始化')
        self.logger.normalInfo('创建索引及选项列表')
        self.parent = parent
        self.__index = []
        self.chooses_cartons = []

    def getModel(self, model, projectVessel):
        self.logger.getFuncName('getModel')
        self.logger.normalInfo('获取目标模型')
        self.__model = model
        self._vessel_project = projectVessel

    # 根据模型机柜类型选择对应的卡件选项
    def getChoosesAndIndex(self, cabinet, dataVessel):
        self.logger.getFuncName('getChoosesAndIndex')
        self.logger.normalInfo('根据机柜类型将model索引加入index列表')
        self.cabinet = cabinet
        if cabinet == '通用':
            self.logger.normalInfo('获取通用卡件种类数据')
            self.chooses_cartons = dataVessel.getCurrentCabinetCartons()
            self.logger.normalInfo('向选项列表内插入空选项')
            self.chooses_cartons.insert(0, 'NULL')
            self.logger.varInfo('模板代理中，选项框的内容->\n' + str(self.chooses_cartons))

            self.logger.normalInfo('进入循环，根据坐标获取各单元格索引并填入列表')
            for num_row in range(0, 10):
                for num_column in range(0, 5):
                    for turn in range(0, 6):
                        index_SL1 = self.__model.index(7 + num_row * 12, 1 + turn + num_column * 6)
                        index_SL2 = self.__model.index(8 + num_row * 12, 1 + turn + num_column * 6)
                        index_SL3 = self.__model.index(9 + num_row * 12, 1 + turn + num_column * 6)
                        index_SL4 = self.__model.index(10 + num_row * 12, 1 + turn + num_column * 6)
                        index_SL5 = self.__model.index(11 + num_row * 12, 1 + turn + num_column * 6)
                        index_SL6 = self.__model.index(12 + num_row * 12, 1 + turn + num_column * 6)
                        index_SL7 = self.__model.index(13 + num_row * 12, 1 + turn + num_column * 6)
                        index_SL8 = self.__model.index(14 + num_row * 12, 1 + turn + num_column * 6)
                        self.__index.append(index_SL1)
                        self.__index.append(index_SL2)
                        self.__index.append(index_SL3)
                        self.__index.append(index_SL4)
                        self.__index.append(index_SL5)
                        self.__index.append(index_SL6)
                        self.__index.append(index_SL7)
                        self.__index.append(index_SL8)

        elif cabinet == '高密度':
            self.logger.normalInfo('获取高密度卡件种类数据')
            self.chooses_cartons = dataVessel.getHighDensityCabinetCartons()
            self.logger.normalInfo('向选项列表内插入空选项')
            self.chooses_cartons.insert(0, 'NULL')
            self.logger.varInfo('模板代理中，选项框的内容->\n'+str(self.chooses_cartons))

            self.logger.normalInfo('进入循环，根据坐标获取各单元格索引并填入列表')
            for num_row in range(0, 10):
                for num_column in range(0, 5):
                    for turn in range(0, 8):
                        index_SL1 = self.__model.index(7 + num_row * 12, 1 + turn + num_column * 8)
                        index_SL2 = self.__model.index(8 + num_row * 12, 1 + turn + num_column * 8)
                        index_SL3 = self.__model.index(9 + num_row * 12, 1 + turn + num_column * 8)
                        index_SL4 = self.__model.index(10 + num_row * 12, 1 + turn + num_column * 8)
                        index_SL5 = self.__model.index(11 + num_row * 12, 1 + turn + num_column * 8)
                        index_SL6 = self.__model.index(12 + num_row * 12, 1 + turn + num_column * 8)
                        index_SL7 = self.__model.index(13 + num_row * 12, 1 + turn + num_column * 8)
                        index_SL8 = self.__model.index(14 + num_row * 12, 1 + turn + num_column * 8)
                        self.__index.append(index_SL1)
                        self.__index.append(index_SL2)
                        self.__index.append(index_SL3)
                        self.__index.append(index_SL4)
                        self.__index.append(index_SL5)
                        self.__index.append(index_SL6)
                        self.__index.append(index_SL7)
                        self.__index.append(index_SL8)
        else:
            pass

        self.logger.normalInfo('创建不含KB4系列继电器的选项容器')
        self.chooses_cartons_delKB = copy.deepcopy(self.chooses_cartons)
        self.logger.varInfo('未删除时的选项容器内容->\n'+str(self.chooses_cartons_delKB))
        num_chooses_cartons_delKB = len(self.chooses_cartons_delKB)
        for index_carton in range(num_chooses_cartons_delKB-1, -1, -1):
            self.logger.normalInfo('第'+str(index_carton)+'个')
            if 'KB' in self.chooses_cartons_delKB[index_carton]:
                self.chooses_cartons_delKB.pop(index_carton)
            else:
                continue

        self.logger.varInfo('无KB继电器选项的容器内容->\n'+str(self.chooses_cartons_delKB))

    # 创建编辑控件
    def createEditor(self, parent, option, index):
        self.logger.getFuncName('createEditor')
        self.logger.normalInfo('根据索引是否在列表范围内，判别代理样式')
        if index in self.__index:
            if self.cabinet == '通用':
                index_modelToChooese = int(self.__index.index(index))
                if (index_modelToChooese % 2) == 0:
                    editor = QComboBox(parent)
                    editor.setEditable(False)
                    editor.addItems(self.chooses_cartons)
                else:
                    editor = QComboBox(parent)
                    editor.setEditable(False)
                    editor.addItems(self.chooses_cartons_delKB)
            else:
                if index.column()%2 == 1:
                    editor = QComboBox(parent)
                    editor.setEditable(False)
                    editor.addItems(self.chooses_cartons)
                else:
                    editor = QComboBox(parent)
                    editor.setEditable(False)
                    editor.addItems(self.chooses_cartons_delKB)
        else:
            editor = QLineEdit(parent)

        return editor

    # 从模型获取数据
    def setEditorData(self, editor, index):
        self.logger.getFuncName('setEditorData')
        self.logger.normalInfo('获取索引对应的model内容')
        model = index.model()
        text = model.data(index, Qt.DisplayRole)
        self.logger.normalInfo('将内容放入编辑器中')
        if index in self.__index:
            editor.setCurrentText(text)
        else:
            editor.setText(text)

    # 将数据从编辑器更新到model
    def setModelData(self, editor, model, index):
        self.logger.getFuncName('setModelData')
        self.logger.normalInfo('获取编辑器中对应索引的内容')
        if index in self.__index:
            text = editor.currentText()
            index_up = model.index(index.row()-1, index.column())
            index_left = model.index(index.row(), index.column()-1)
            data_up = model.data(index_up, Qt.DisplayRole)
            data_left = model.data(index_left, Qt.DisplayRole)
            if data_up is None:
                data_up = ''
            elif data_left is None:
                data_left = ''
            self.logger.normalInfo('进入判别')
            try:
                if text == 'NULL':
                    self.logger.normalInfo('进入判别->NULL')
                    text = ''
                    self.logger.normalInfo('NULL->判别结束')
                elif 'KB' in text and self.cabinet == '通用':
                    self.logger.normalInfo('进入判别->通用KB')
                    index_down = model.index(index.row() + 1, index.column())
                    model.setData(index_down, '', Qt.DisplayRole)
                    self.logger.normalInfo('通用KB->判别结束')
                elif 'KB' in text and self.cabinet == '高密度':
                    self.logger.normalInfo('进入判别->高密度KB')
                    index_right = model.index(index.row(), index.column() + 1)
                    model.setData(index_right, '', Qt.DisplayRole)
                    self.logger.normalInfo('高密度KB->判别结束')
                elif 'KB' in data_up and self.cabinet == '通用':
                    self.logger.normalInfo('进入判别->上方KB')
                    dlg = QMessageBox(self.parent)
                    dlg.setIcon(QMessageBox.Critical)
                    dlg.setText('上方存在继电器')
                    dlg.show()
                    text = ''
                    self.logger.normalInfo('上方KB->判别结束')
                elif 'KB' in data_left and self.cabinet == '高密度':
                    self.logger.normalInfo('进入判别->左方KB')
                    dlg = QMessageBox(self.parent)
                    dlg.setIcon(QMessageBox.Critical)
                    dlg.setText('左方存在继电器')
                    dlg.show()
                    text = ''
                    self.logger.normalInfo('左方KB->判别结束')
            except:
                self.logger.tranceInfo()

            self.logger.normalInfo('结束判别')
        else:
            text = editor.text()

        self.logger.normalInfo('将内容输入到model中')
        self._vessel_project.project[9] = False
        model.setData(index, text, Qt.DisplayRole)

    # 设置控件大小
    def updateEditorGeometry(self, editor, option, index):
        self.logger.getFuncName('updateEditorGeometry')
        editor.setGeometry(option.rect)


class PotsIoFormViewDelegate(QStyledItemDelegate):
    def __init__(self):
        super().__init__()
        # 默认、DCS——2线制
        # 外部供电——4线制
        # 220V——电压接线
        #
        self.logger = PotsLogger('PotsIoFormViewDelegate')
        self.logger.normalInfo('进入IO清册代理并创建选项内容')


    def getModel(self, model, projectVessel):
        self.logger.getFuncName('getModel')
        self.logger.normalInfo('获取目标模型')
        self.__model = model
        self._vessel_project = projectVessel

    # 根据模型机柜类型选择对应的卡件选项
    def getChoosesAndIndex(self, dataVessel):
        self.logger.getFuncName('getChoosesAndIndex')
        self.logger.normalInfo('获取数据库容器')
        self.__dataVessel = dataVessel

    # 创建编辑控件
    def createEditor(self, parent, option, index):
        self.logger.getFuncName('createEditor')
        self.logger.normalInfo('判别编辑器设置依据')
        model = index.model()
        col = index.column()

        list_ediableColumn = [10, 11, 12, 13, 14, 15, 16, 18, 19, 20, 21, 22, 23, 24,
                              25, 26, 27, 28, 29, 30, 31]
        try:
            if col in list_ediableColumn and index.row() >= 2:
                editor = QLineEdit(parent)
            elif col == 17 and index.row() >= 3:
                index_carton = model.index(index.row(), 1)
                carton = index_carton.data(Qt.DisplayRole)
                editor = QComboBox(parent)
                editor.setEditable(False)
                chooses = self.__dataVessel.getWires(carton)
                editor.addItems(chooses)
            else:
                editor = QLineEdit(parent)
                editor.setReadOnly(True)
        except:
            self.logger.tranceInfo()

        return editor

    # 从模型获取数据
    def setEditorData(self, editor, index):
        self.logger.getFuncName('setEditorData')
        model = index.model()
        self.logger.normalInfo('获取索引对应的model内容')
        text = model.data(index, Qt.DisplayRole)
        col = index.column()
        self.logger.normalInfo('将内容放入编辑器中')
        try:
            if col == 17 and index.row() >= 3:
                editor.setCurrentText(text)
            else:
                editor.setText(text)
        except:
            self.logger.tranceInfo()

    # 将数据从编辑器更新到model
    def setModelData(self, editor, model, index):
        self.logger.getFuncName('setModelData')
        col = index.column()
        if col == 17 and index.row() >= 3:
            self.logger.normalInfo('获取编辑器中对应索引的内容')
            text = editor.currentText()
            self.logger.normalInfo('获取当前索引所在行')
            row = index.row()

            self.logger.normalInfo('获取当前行所对应卡件及其通道号')
            index_cartonName = model.index(row, 1)
            index_roadNumFore = model.index(row, 8)
            name_carton = model.data(index_cartonName, Qt.DisplayRole)
            num_roadFore = model.data(index_roadNumFore, Qt.DisplayRole)
            self.logger.normalInfo('根据通道号获取卡件相应属性')
            try:
                if num_roadFore[-2] == '0':
                    num_road = num_roadFore[-1]
                else:
                    num_road = num_roadFore[-2:]
                data_carton = self.__dataVessel.getThoroughfare(name_carton, num_road, text)
                if data_carton == []:
                    self.logger.abnormalInfo('该线制不存在')
                    self.logger.normalInfo('获取编辑器中对应索引的内容')
                    text = editor.text()
                else:
                    self.logger.normalInfo('设置model内容')
                    carton_C4 = QStandardItem(data_carton[0])
                    carton_C5 = QStandardItem(data_carton[1])
                    carton_C6 = QStandardItem(data_carton[2])
                    carton_C7 = QStandardItem(data_carton[3])
                    model.setItem(row, 4, carton_C4)
                    model.setItem(row, 5, carton_C5)
                    model.setItem(row, 6, carton_C6)
                    model.setItem(row, 7, carton_C7)
            except:
                self.logger.tranceInfo()
        else:
            self.logger.normalInfo('获取编辑器中对应索引的内容')
            text = editor.text()

        self._vessel_project.project[9] = False
        model.setData(index, text, Qt.DisplayRole)

    # 设置控件大小
    def updateEditorGeometry(self, editor, option, index):
        self.logger.getFuncName('updateEditorGeometry')
        editor.setGeometry(option.rect)

#测试函数
def test():
    pass

#开启动作
if __name__ == '__main__':
    test()
