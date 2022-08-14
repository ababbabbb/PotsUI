# python包
from PyQt5.QtWidgets import QMdiArea, QMdiSubWindow, QMessageBox, QLabel
from PyQt5.QtCore import Qt

# 自建包
from CodeTools.loggingTool import PotsLogger

#类
class TabModeSubWindow(QMdiSubWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.logger = PotsLogger('TabModeSubWindow')
        self.setAttribute(Qt.WA_DeleteOnClose)

    def getMdi(self, mdiObj):
        self.logger.getFuncName('getMdi')
        self.mdi = mdiObj

    def closeEvent(self, event):
        self.logger.getFuncName('closeEvent')
        self.logger.normalInfo('子窗体关闭')
        # self.mdi.removeSubWindow(self)
        lable_deleted = QLabel('')
        self.setWidget(lable_deleted)
        event.accept()

class PotsMdiWindow():
    def __init__(self, parent=None, tree=None):
        super().__init__()
        self.logger = PotsLogger('PotsMdiWindow')
        self.logger.normalInfo('初始化mdi窗口')
        self.mdi = QMdiArea(parent)
        self.mdi.setViewMode(1)
        self.mdi.setTabsClosable(True)
        self.mdi.setTabsMovable(True)
        self.treeObject = tree

    def addMdiWidgets(self, widget, name):
        self.logger.getFuncName('addMdiWidgets')
        self.logger.normalInfo('匹配当前活跃子窗体')
        list_subWindows = self.mdi.subWindowList()
        self.logger.varInfo('这是添加subuwindow时，开始时获取的子窗体列表->\n'+str(list_subWindows))
        for subWindow in list_subWindows:
            title_subWindow = subWindow.windowTitle()
            self.logger.varInfo('子窗体标题有->\n'+str(title_subWindow))
            # if title_subWindow == name and subWindow == self.mdi.activeSubWindow():
            if title_subWindow == name:
                self.logger.abnormalInfo('所添加的窗口已存在')
                dlg = QMessageBox(self.mdi)
                dlg.setIcon(QMessageBox.Critical)
                dlg.setText('当前窗口已打开')
                dlg.show()
                return 0
        self.logger.normalInfo('创建mdi窗口的子窗口，并放入mdi区域')
        self.subWindow_new = TabModeSubWindow(parent=self.mdi)
        self.subWindow_new.getMdi(self.mdi)
        self.subWindow_new.setWindowTitle(name)
        self.logger.normalInfo('当前名称为->'+str(name))
        self.logger.varInfo('当前控件内容->\n'+str(widget))
        try:
            self.logger.normalInfo('添加控件到子窗体')
            self.subWindow_new.setWidget(widget)
        except:
            self.logger.tranceInfo()
            return 0
        # self.subWindow_new.setAttribute(Qt.WA_DeleteOnClose)
        self.logger.normalInfo('subwindow创建并样式设置完毕')
        self.logger.normalInfo('将子窗体放入mdi停靠区')
        self.mdi.addSubWindow(self.subWindow_new)

        self.logger.normalInfo('创建并添加完毕')

        self.logger.varInfo('当前mdi区域包含的窗体->\n'+str(self.mdi.subWindowList()))
        self.logger.varInfo('窗体的名字->\n'+str(name))

        self.subWindow_new.show()

    def closeAll(self):
        self.logger.getFuncName('closeAll')
        self.logger.normalInfo('关闭所有窗体')
        self.mdi.closeAllSubWindows()

    def getCurrentWidgets(self):
        self.logger.getFuncName('getCurrentWidgets')
        try:
            self.logger.normalInfo('获取当前激活的子窗体')
            currentSubwindow = self.mdi.activeSubWindow()
            name_view = currentSubwindow.windowTitle()
            view = currentSubwindow.widget()
        except Exception as e:
            self.logger.warringInfo(str(e))
            return None, None
        else:
            return view, name_view


#测试函数
def test():
    pass

#开启动作
if __name__ == '__main__':
    test()
