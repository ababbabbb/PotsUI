# python包
from PyQt5.QtWidgets import QTreeWidget, QTreeWidgetItem

# 自建包
from CodeTools.loggingTool import PotsLogger

#类
class PotsTree():
    def __init__(self, parent=None, allVessel=None):
        self.parent = parent
        self.logger = PotsLogger('PotsTree')
        self.logger.normalInfo('树节点对象初始化')

        self.vessel_all = allVessel

        self.logger.normalInfo('进入自检')
        self.checkSelf()

    def getInfo(self, cabinet=None, vessel=None, mdiWindow=None):
        self.logger.getFuncName('getInfo')
        self.logger.normalInfo('获取树节点信息')
        self.cabinet = cabinet
        self.vessel = vessel
        self.mdiWindow = mdiWindow

    def checkSelf(self):
        self.logger.getFuncName('checkSelf')

        self.tree = QTreeWidget(self.parent)
        self.tree.setColumnCount(1)
        self.tree.setHeaderLabel('工程面板')
        self.tree.itemDoubleClicked.connect(self.respondAction)

    def addProject(self, projectName):
        self.logger.getFuncName('addProject')
        self.logger.normalInfo('进入项目树节点构造')
        self.projectName = projectName
        # 一级节点
        self.logger.normalInfo('构造一级节点')
        item_first = QTreeWidgetItem()
        item_first.setText(0, '项目_' + self.projectName)
        self.tree.addTopLevelItem(item_first)
        self.insertProjectItemToVessel(item_first)

        self.logger.normalInfo('构造二级节点')
        # 二级节点——model
        item_model = QTreeWidgetItem()
        item_model.setText(0, 'startForm')
        item_first.addChild(item_model)
        item_modelForm = QTreeWidgetItem()
        item_modelForm.setText(0, '模板_'+self.projectName)
        item_model.addChild(item_modelForm)
        self.insertmodelTreeItemToVessel(item_modelForm)

        # 二级节点——IO
        item_IO = QTreeWidgetItem()
        item_IO.setText(0, 'IOForm')
        item_first.addChild(item_IO)
        for i in range(1, 51):
            item_modelForm = QTreeWidgetItem()
            item_modelForm.setText(0, '清册' + str(i).zfill(2) + self.projectName)
            item_IO.addChild(item_modelForm)
            self.insertIOTreeItemInsert(item_modelForm, i)

    def insertProjectItemToVessel(self, item):
        self.logger.getFuncName('insertProjectItemToVessel')
        self.logger.normalInfo('创建项目树节点属性')
        property_item = {
            'parent': self.parent,
            'type': 'item_project',
            'name': '项目_' + self.projectName,
            'cabinet': self.cabinet
        }
        self.vessel.projectTreeItemInsert(item, property_item)

        self.logger.varInfo('项目节点变化')
        self.logger.varInfo(str(self.vessel.project))

    def insertmodelTreeItemToVessel(self, item):
        self.logger.getFuncName('insertProjectItemToVessel')
        self.logger.normalInfo('创建model节点属性')
        property_item = {
            'parent': self.parent,
            'type': 'item_model',
            'name': '模板_' + self.projectName,
            'cabinet': self.cabinet
        }
        self.vessel.modelTreeItemInsert(item, property_item)

        self.logger.varInfo('model节点变化')
        self.logger.varInfo(str(self.vessel.project))

    def insertIOTreeItemInsert(self, item, i):
        self.logger.getFuncName('insertProjectItemToVessel')
        self.logger.normalInfo('创建清册树节点属性')
        property_item = {
            'parent': self.parent,
            'type': 'item_IO',
            'name': '清册' + str(i).zfill(2) + self.projectName,
            'cabinet': self.cabinet
        }
        self.vessel.IOTreeItemInsert(item, property_item)

        self.logger.varInfo('IO节点变化')
        self.logger.varInfo(str(self.vessel.project))

    def respondAction(self):
        self.logger.getFuncName('respondAction')
        self.logger.normalInfo('进入节点响应逻辑')
        self.logger.normalInfo('获取当前受点击节点名称')
        item_current = self.tree.currentItem()
        name_currentItem = item_current.text(0)

        if '模板' in name_currentItem:
            name_project = name_currentItem[3:]
            self.logger.varInfo('此时项目名称为->\n' + str(name_project))
        elif '清册' in name_currentItem:
            name_project = name_currentItem[4:]
            self.logger.varInfo('此时项目名称为->\n' + str(name_project))
        else:
            return 0

        self.logger.varInfo('此时选中的树节点的名称为->\n'+str(name_currentItem))
        for vessel in self.vessel_all:
            if name_project == vessel[0]:
                vessel_project = vessel
                break
            else:
                continue

        self.logger.normalInfo('根据获得的名称进入mdi区域子窗口添加逻辑')
        if '模板' in name_currentItem:
            for v in vessel_project[6]:
                if name_currentItem == v[1]['name']:
                    self.mdiWindow.addMdiWidgets(v[0], name_currentItem)
        elif '清册' in name_currentItem:
            for v in vessel_project[7]:
                if name_currentItem == v[1]['name']:
                    self.mdiWindow.addMdiWidgets(v[0], name_currentItem)
        else:
            pass

#测试函数
def test():
    pass

#开启动作
if __name__ == '__main__':
    test()
