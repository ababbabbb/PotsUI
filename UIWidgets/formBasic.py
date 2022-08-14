# python包
from abc import abstractmethod, ABCMeta
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QStandardItemModel, QStandardItem
from PyQt5.QtWidgets import QMessageBox

# 自建包

#类
class PotsModel(QStandardItemModel):

    def data(self, index, role=None):
        if role == Qt.TextAlignmentRole:
            return Qt.AlignCenter
        return QStandardItemModel.data(self, index, role)


class PotsFormCreatorBasic(metaclass=ABCMeta):

    def __init__(self, parent=None, modelType=None, modelName=None):
        self.parent = parent
        self.modelType(modelType)
        self.name_model = modelName

    def insertValue(self, target, row, col, value):
        item = QStandardItem(value)
        target.setItem(row, col, item)

    # 信息提示对话框
    def tipsDiag(self, strInfo):
        # 提示信息
        dlg = QMessageBox(self)
        dlg.setIcon(QMessageBox.Critical)
        dlg.setText('来自'+str(self.name_model)+'的消息:'+strInfo)
        dlg.show()

    # 定义model类型——startForm/IOform/matchinForm/analysForm/...
    def modelType(self, type=None):
        self.type_model = type

    # 校验自身属性正确性
    @abstractmethod
    def checkSelf(self):
        pass

    # 将自身放入管理类内
    @abstractmethod
    def insertModelToVessel(self, model):
        pass

    @abstractmethod
    def insertViewToVessel(self, view):
        pass

    @abstractmethod
    def createModel(self):
        pass

    @abstractmethod
    def createView(self):
        pass


#测试函数
def test():
    pass

#开启动作
if __name__ == '__main__':
    test()
