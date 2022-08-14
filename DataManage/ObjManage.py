# python包
from PyQt5.QtWidgets import QMessageBox, QTableView
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QStandardItem
import re

# 自建包
from CodeTools.loggingTool import PotsLogger

# 类
class PotsGlobalVessel():

    def __init__(self):
        self.logger = PotsLogger('PotsGlobalVessel')
        self.logger.normalInfo('创建总容器')
        self.globalVessel = []


class PotsVessel():
    def __init__(self, projectName):
        self.logger = PotsLogger('PotsVessel')
        self.logger.normalInfo('初始化项目容器对象，并创建各类容器')
        self.project = [projectName]    # 0

        self.projectTreeItem = []   # 1
        self.modelTreeItem = []    # 2
        self.IOTreeItem = []    # 3

        self.modelVessel_start = []     # 4
        self.modelVessel_IO = []    # 5

        self.viewVessel_start = []      # 6
        self.viewVessel_IO = []     # 7

        self.projectPath = None
        property_path = {
            'name': projectName
        }
        self.pathVessel = []    # 8
        self.pathInsert(self.projectPath, property_path)

        self.savedKeyWord = False   # 9

        self.project.append(self.projectTreeItem)
        self.project.append(self.modelTreeItem)
        self.project.append(self.IOTreeItem)
        self.project.append(self.modelVessel_start)
        self.project.append(self.modelVessel_IO)
        self.project.append(self.viewVessel_start)
        self.project.append(self.viewVessel_IO)
        self.project.append(self.pathVessel)
        self.project.append(self.savedKeyWord)

        self.logger.varInfo('第一次查看项目容器内容')
        self.logger.varInfo(str(self.project))

    def insert(self, secondaryVessel, vessel):
        self.logger.getFuncName('insert')
        self.logger.normalInfo('容器添加内容')
        vessel.append(secondaryVessel)
        self.logger.normalInfo('添加结束')

    def startModelInsert(self, model, property_m):
        self.logger.getFuncName('startModelInsert')
        secondaryVessel = []
        secondaryVessel.append(model)
        secondaryVessel.append(property_m)
        self.insert(secondaryVessel, self.modelVessel_start)

    def startViewInsert(self, view, property_v):
        self.logger.getFuncName('startViewInsert')
        secondaryVessel = []
        secondaryVessel.append(view)
        secondaryVessel.append(property_v)
        self.insert(secondaryVessel, self.viewVessel_start)

    def IOmodelInsert(self, model, property_m):
        self.logger.getFuncName('IOmodelInsert')
        secondaryVessel = []
        secondaryVessel.append(model)
        secondaryVessel.append(property_m)
        self.insert(secondaryVessel, self.modelVessel_IO)

    def IOViewInsert(self, view, property_v):
        self.logger.getFuncName('IOViewInsert')
        secondaryVessel = []
        secondaryVessel.append(view)
        secondaryVessel.append(property_v)
        self.insert(secondaryVessel, self.viewVessel_IO)

    def projectTreeItemInsert(self, item, property_i):
        self.logger.getFuncName('projectTreeItemInsert')
        secondaryVessel = []
        secondaryVessel.append(item)
        secondaryVessel.append(property_i)
        self.insert(secondaryVessel, self.projectTreeItem)

    def modelTreeItemInsert(self, item, property_i):
        self.logger.getFuncName('modelTreeItemInsert')
        secondaryVessel = []
        secondaryVessel.append(item)
        secondaryVessel.append(property_i)
        self.insert(secondaryVessel, self.modelTreeItem)

    def IOTreeItemInsert(self, item, property_i):
        self.logger.getFuncName('IOTreeItemInsert')
        secondaryVessel = []
        secondaryVessel.append(item)
        secondaryVessel.append(property_i)
        self.insert(secondaryVessel, self.IOTreeItem)

    def pathInsert(self, path, property_p):
        self.logger.getFuncName('pathInsert')
        secondaryVessel = []
        secondaryVessel.append(path)
        secondaryVessel.append(property_p)
        self.insert(secondaryVessel, self.pathVessel)


class PotsConfirmForms():
    def __init__(self, projectVessel=None, dataVessel=None, mainWindow=None, treeObj=None, mdiObj=None, cover=None):
        self.logger = PotsLogger('PotsConfirmForms')
        self.logger.normalInfo('表单连接对象初始化')

        self.logger.normalInfo('传入连接对象')
        self.coverWindow = cover
        self.projectVessel = projectVessel
        self.dataVessel = dataVessel
        self.mainWindow = mainWindow
        self.treeObj = treeObj
        self.mdiObj = mdiObj
        self.cabinet = projectVessel[4][0][1]['cabinet']

        self.logger.normalInfo('进入自检')
        self.checkSelf()

    def checkSelf(self):
        self.logger.getFuncName('checkSelf')
        if self.projectVessel is None:
            self.logger.abnormalInfo('无项目容器传入表单连接器')
            dlg = QMessageBox(self.mainWindow)
            dlg.setIcon(QMessageBox.Critical)
            dlg.setText('无项目容器传入表单连接器')
            dlg.show()
        else:
            self.logger.normalInfo('进入连接')
            self.confirmFunction()

    def confirmFunction(self):
        self.logger.getFuncName('confirmFunction')
        self.logger.normalInfo('关闭当前全部窗口')
        self.mdiObj.closeAll()

        self.logger.normalInfo('设置进度为28')
        self.coverWindow.setValue(28)

        self.logger.normalInfo('获取模件model及IO视图')
        self.startModle = self.projectVessel[4][0][0]
        self.IOViewVessel = self.projectVessel[7]

        self.logger.normalInfo('设置进度为40')
        self.coverWindow.setValue(40)

        # 添加IO行数判断，倘若行数超过2，清空剩余全部行再继续执行
        self.logger.normalInfo('进入遍历，根据模件model更改IO视图内容')
        value_add = 0
        for viewList in self.IOViewVessel:
            value_add += 1.2
            IOViewInVessel = viewList[0]
            IOModleInVessel = viewList[0].model()
            self.logger.normalInfo('清除IO清册超过2行的内容')
            IOModleRow_current = IOModleInVessel.rowCount()
            self.logger.varInfo('创建清册时，清册本身model行数'+str(IOModleRow_current))
            if IOModleRow_current > 2:
                for i in range(IOModleRow_current, 1, -1):
                    index_deleted = IOModleInVessel.index(i, 1)
                    self.logger.varInfo('此时删除的清册原有行-> ' + str(i))
                    self.logger.varInfo('行数为->' + str(index_deleted.row()))
                    IOModleInVessel.removeRow(index_deleted.row())
            confirmIndex = viewList[1]['startFormSheetsIndex']
            cabinet = viewList[1]['cabinet']
            viewList[1]['mergeList'] = []
            mergeList = viewList[1]['mergeList']
            self.logger.normalInfo('根据内容，更新清册内容')
            self.addIOPart(self.startModle, IOModleInVessel, confirmIndex, cabinet, mergeList)

            self.logger.normalInfo('设置进度为' + str(value_add / 2 + 50))
            self.coverWindow.setValue(value_add / 2 + 50)

            self.logger.normalInfo('进入遍历，合并单元格')
            for coordinate_merge in mergeList:
                IOViewInVessel.setSpan(coordinate_merge[0],
                                       coordinate_merge[1],
                                       coordinate_merge[2],
                                       coordinate_merge[3])

            self.logger.normalInfo('设置进度为' + str(value_add/2 + 50))
            self.coverWindow.setValue(value_add/2 + 50)

    def addIOPart(self, startModle, IOModle, index, cabinet, mergeList):
        self.logger.getFuncName('addIOPart')
        num_cartonsIndex = 1  # 清册有效行计数
        self.logger.normalInfo('区别机柜类型')
        if cabinet == '高密度':
            columnNum_sheet = 8
        elif cabinet == '通用':
            columnNum_sheet = 6

        self.logger.normalInfo('根据模件表单号确定待生成内容')
        col_sheet = (index % 5) - 1
        row_sheet = index // 5

        self.logger.normalInfo('获取对应表单的DPU及机柜名')
        index_DPU = startModle.index(4 + row_sheet * 12, 1 + col_sheet * columnNum_sheet)
        index_CAB = startModle.index(5 + row_sheet * 12, 1 + col_sheet * columnNum_sheet)

        value_DPU = startModle.data(index_DPU, Qt.DisplayRole)
        value_CAB = startModle.data(index_CAB, Qt.DisplayRole)

        if value_DPU is None or value_CAB is None:
            self.logger.normalInfo('当前表单依据DPU、CAB判别为无内容')
            return 0
        value_DPU = re.sub('[\u4e00-\u9fa5]', '', value_DPU)
        self.logger.normalInfo('进入循环获取模板表单每一列的值，并填入对应IO清册')
        for turn in range(0, columnNum_sheet):
            index_BRC = startModle.index(6 + row_sheet * 12, 1 + turn + col_sheet * columnNum_sheet)

            value_BRC = startModle.data(index_BRC, Qt.DisplayRole)
            if value_BRC is None or value_BRC == '':
                self.logger.normalInfo('依据BRC判别当前列无内容')
                continue
            self.logger.normalInfo('1~7列索引')
            index_SL1 = startModle.index(7 + row_sheet * 12, 1 + turn + col_sheet * columnNum_sheet)
            index_SL2 = startModle.index(8 + row_sheet * 12, 1 + turn + col_sheet * columnNum_sheet)
            index_SL3 = startModle.index(9 + row_sheet * 12, 1 + turn + col_sheet * columnNum_sheet)
            index_SL4 = startModle.index(10 + row_sheet * 12, 1 + turn + col_sheet * columnNum_sheet)
            index_SL5 = startModle.index(11 + row_sheet * 12, 1 + turn + col_sheet * columnNum_sheet)
            index_SL6 = startModle.index(12 + row_sheet * 12, 1 + turn + col_sheet * columnNum_sheet)
            index_SL7 = startModle.index(13 + row_sheet * 12, 1 + turn + col_sheet * columnNum_sheet)
            index_SL8 = startModle.index(14 + row_sheet * 12, 1 + turn + col_sheet * columnNum_sheet)

            # 添加标题
            self.logger.normalInfo('添加标题')
            IoModleRow_current = IOModle.rowCount()
            part_C1 = QStandardItem('Branch_' + value_BRC)
            part_C7 = QStandardItem(value_DPU)
            IOModle.setItem(IoModleRow_current, 1, part_C1)
            IOModle.setItem(IoModleRow_current, 8, part_C7)

            self.logger.normalInfo('将模件对应表单1~7列内容放入列表待填入清册')
            value_SL1 = startModle.data(index_SL1, Qt.DisplayRole)
            value_SL2 = startModle.data(index_SL2, Qt.DisplayRole)
            self.logger.normalInfo('当前行对应卡件为->'+str(value_SL2))
            value_SL3 = startModle.data(index_SL3, Qt.DisplayRole)
            value_SL4 = startModle.data(index_SL4, Qt.DisplayRole)
            value_SL5 = startModle.data(index_SL5, Qt.DisplayRole)
            value_SL6 = startModle.data(index_SL6, Qt.DisplayRole)
            value_SL7 = startModle.data(index_SL7, Qt.DisplayRole)
            value_SL8 = startModle.data(index_SL8, Qt.DisplayRole)
            list_startSheetColumnValue = []
            list_startSheetColumnValue.append(value_SL1)
            list_startSheetColumnValue.append(value_SL2)
            list_startSheetColumnValue.append(value_SL3)
            list_startSheetColumnValue.append(value_SL4)
            list_startSheetColumnValue.append(value_SL5)
            list_startSheetColumnValue.append(value_SL6)
            list_startSheetColumnValue.append(value_SL7)
            list_startSheetColumnValue.append(value_SL8)

            self.logger.normalInfo('填入清册内容')
            for num in range(3, 11):
                cartonInStartForm = list_startSheetColumnValue[num-3]
                IoModleRow_current = IOModle.rowCount()
                if cartonInStartForm is not None and cartonInStartForm != '' and cartonInStartForm != 'None':
                    try:
                        data_carton = self.dataVessel.getTerminalData(cartonInStartForm)
                    except:
                        self.logger.tranceInfo()
                    num_cartonData = len(data_carton)
                    mergeList_C1 = [IoModleRow_current, 1, num_cartonData, 1]
                    mergeList_C2 = [IoModleRow_current, 2, num_cartonData, 1]
                    mergeList_C3 = [IoModleRow_current, 3, num_cartonData, 1]
                    mergeList.append(mergeList_C1)
                    mergeList.append(mergeList_C2)
                    mergeList.append(mergeList_C3)
                    for n in range(0, num_cartonData):
                        self.logger.normalInfo('判别行内容')
                        if 'KB431A' in cartonInStartForm or 'KB432C' in cartonInStartForm or \
                                'KB432D' in cartonInStartForm or 'KB432A' in cartonInStartForm:
                            carton_C0 = QStandardItem('')
                            carton_C1 = QStandardItem(cartonInStartForm)
                            carton_C2 = QStandardItem(value_CAB)
                            carton_C3 = QStandardItem(value_BRC + str(num - 2).zfill(2) + '/'
                                                      + str(num - 1))
                            carton_C4 = QStandardItem(data_carton[n][4])
                            carton_C5 = QStandardItem(data_carton[n][5])
                            carton_C6 = QStandardItem(data_carton[n][6])
                            carton_C7 = QStandardItem(data_carton[n][7])
                            if self.cabinet == '通用':
                                if n < num_cartonData / 2:
                                    carton_C8 = QStandardItem('上' +
                                                              value_BRC + str(n+1).zfill(2))
                                    carton_C9 = QStandardItem('V4::DPU' + value_DPU + '.HW.' +
                                                      data_carton[n][1] +
                                                      value_BRC + str(num - 2).zfill(2) + str(n+1).zfill(2)
                                                      )
                                else:
                                    carton_C8 = QStandardItem('下' +
                                                              value_BRC + str(n+1).zfill(2))
                                    carton_C9 = QStandardItem('V4::DPU' + value_DPU + '.HW.' +
                                                      data_carton[n][1] +
                                                      value_BRC + str(num - 2).zfill(2) + str(n+1).zfill(2)
                                                      )
                            else:
                                if n < num_cartonData / 2:
                                    carton_C8 = QStandardItem('左' +
                                                              str(int(value_BRC)).zfill(2) + str(n+1).zfill(2))
                                    carton_C9 = QStandardItem('V4::DPU' + value_DPU + '.HW.' +
                                                      data_carton[n][1] +
                                                      str(int(value_BRC)).zfill(2) + str(num - 2).zfill(2) + str(n + 1).zfill(2))
                                else:
                                    carton_C8 = QStandardItem('右' +
                                                              str(int(value_BRC)+1).zfill(2) + str(n+1).zfill(2))
                                    carton_C9 = QStandardItem('V4::DPU' + value_DPU + '.HW.' +
                                                      data_carton[n][1] +
                                                      str(int(value_BRC)+1).zfill(2) + str(num - 2).zfill(2) + str(n+1).zfill(2))
                            carton_C18 = QStandardItem(data_carton[n][2])
                            IOModle.setItem(IoModleRow_current + n, 0, carton_C0)
                            IOModle.setItem(IoModleRow_current + n, 1, carton_C1)
                            IOModle.setItem(IoModleRow_current + n, 2, carton_C2)
                            IOModle.setItem(IoModleRow_current + n, 3, carton_C3)
                            IOModle.setItem(IoModleRow_current + n, 4, carton_C4)
                            IOModle.setItem(IoModleRow_current + n, 5, carton_C5)
                            IOModle.setItem(IoModleRow_current + n, 6, carton_C6)
                            IOModle.setItem(IoModleRow_current + n, 7, carton_C7)
                            IOModle.setItem(IoModleRow_current + n, 8, carton_C8)
                            IOModle.setItem(IoModleRow_current + n, 9, carton_C9)
                            IOModle.setItem(IoModleRow_current + n, 17, carton_C18)
                        else:
                            carton_C0 = QStandardItem(str(num_cartonsIndex))
                            carton_C1 = QStandardItem(cartonInStartForm)
                            carton_C2 = QStandardItem(value_CAB)
                            carton_C3 = QStandardItem(value_BRC + str(num - 2).zfill(2))
                            carton_C4 = QStandardItem(data_carton[n][4])
                            carton_C5 = QStandardItem(data_carton[n][5])
                            carton_C6 = QStandardItem(data_carton[n][6])
                            carton_C7 = QStandardItem(data_carton[n][7])
                            carton_C8 = QStandardItem(data_carton[n][1] +
                                                      value_BRC + str(num - 2).zfill(2) + str(n + 1).zfill(2))
                            carton_C9 = QStandardItem('V4::DPU' + value_DPU + '.HW.' +
                                                      data_carton[n][1] +
                                                      value_BRC + str(num - 2).zfill(2) + str(n + 1).zfill(2)
                                                      )
                            carton_C18 = QStandardItem(data_carton[n][2])

                            IOModle.setItem(IoModleRow_current + n, 0, carton_C0)
                            IOModle.setItem(IoModleRow_current + n, 1, carton_C1)
                            IOModle.setItem(IoModleRow_current + n, 2, carton_C2)
                            IOModle.setItem(IoModleRow_current + n, 3, carton_C3)
                            IOModle.setItem(IoModleRow_current + n, 4, carton_C4)
                            IOModle.setItem(IoModleRow_current + n, 5, carton_C5)
                            IOModle.setItem(IoModleRow_current + n, 6, carton_C6)
                            IOModle.setItem(IoModleRow_current + n, 7, carton_C7)
                            IOModle.setItem(IoModleRow_current + n, 8, carton_C8)
                            IOModle.setItem(IoModleRow_current + n, 9, carton_C9)
                            IOModle.setItem(IoModleRow_current + n, 17, carton_C18)

                            num_cartonsIndex += 1
                else:
                    continue


#测试函数
def test():
    pass

#开启动作
if __name__ == '__main__':
    test()
