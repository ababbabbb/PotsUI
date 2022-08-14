# python包
from PyQt5.QtWidgets import QDialog, QLabel, QHBoxLayout, QVBoxLayout, QProgressBar, QApplication
from PyQt5.QtCore import Qt, QTimer

# 自建包
from CodeTools.loggingTool import PotsLogger

#类
class PotsCoverWindow(QDialog):
    def __init__(self, parent=None):
        super(PotsCoverWindow, self).__init__(parent)
        self.logger = PotsLogger('PotsCoverWindow')
        self.logger.normalInfo('设置窗体样式')
        self.mainWindow = parent
        self.setWindowTitle('')
        self.setWindowModality(Qt.ApplicationModal)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        # self.setAttribute(Qt.WA_TranslucentBackground)
        self.resize(450, 100)

        self.logger.normalInfo('构造掩盖窗体控件')

        self.label = QLabel('功能运行中，点击将会白屏，不影响使用，但仍请勿点击耐心等待...')

        self.featProgressBar = QProgressBar(self)
        self.featProgressBar.setMinimum(0)
        self.featProgressBar.setMaximum(100)
        self.featProgressBar.setValue(0)

        self.timer = QTimer()
        self.timer.start(100)
        self.timer.timeout.connect(self.updateMainUI)

        labelLayout = QHBoxLayout()
        labelLayout.addWidget(self.label)

        processBarLayout = QHBoxLayout()
        processBarLayout.addWidget(self.featProgressBar)

        layout = QVBoxLayout()
        layout.addLayout(labelLayout)
        layout.addLayout(processBarLayout)
        self.setLayout(layout)

        self.logger.normalInfo('展示掩盖窗体')
        self.show()

    def updateMainUI(self):
        QApplication.processEvents()

    def setValue(self, value):
        self.featProgressBar.setValue(value)

#测试函数
def test():
    pass

#开启动作
if __name__ == '__main__':
    test()
