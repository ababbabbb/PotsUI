# python包
import os
import sys
import time
from pathlib import Path

# 自建包

# 类
class PotsLogFolder:
    path_exe = os.getcwd()
    topFolder_log = path_exe + '\\PotsCellogs'
    folder_log = topFolder_log + time.strftime('\\%Y-%m-%d,%H-%M-%S')
    if os.path.exists(topFolder_log):
        pass
    else:
        os.mkdir(topFolder_log)

    if os.path.exists(folder_log):
        pass
    else:
        os.mkdir(folder_log)

class PotsLogFile:
    flie_bessineseLog = PotsLogFolder.folder_log + '\\bessibese.txt'
    file_varLog = PotsLogFolder.folder_log + '\\varLog.txt'

class PotsLogger(object):
    def __init__(self,  name_class):
        self.path_bessineseLog = PotsLogFile.flie_bessineseLog
        self.path_varLog = PotsLogFile.file_varLog
        self.name_class = name_class
        self.running(self.name_class)

    def loggerWork(self, flag='fb'):
        def work(func):
            def wrapper(any, str_info):
                if 'debug' in Path(__file__).stem:
                    info = func(any, str_info)
                    if flag == 'fb':
                        file = open(PotsLogFile.flie_bessineseLog, 'a+')
                    else:
                        file = open(PotsLogFile.file_varLog, 'a+')
                    print(info, file=file)
                    file.close()
                else:
                    pass
            return wrapper
        return work

    @loggerWork('fb')
    def running(self, name_class):
        info = time.strftime('%H:%M:%S -> ') + 'class ->' + name_class + ' is running'
        return info

    @loggerWork('fb')
    def getFuncName(self, str_info):
        info = time.strftime('%H:%M:%S ->') + self.name_class + '|func ->'+ str_info +' is running'
        return info

    @loggerWork('fb')
    def normalInfo(self, str_info):
        info = time.strftime('%H:%M:%S ->') + self.name_class + '|正常 —>\n' + str_info
        return info

    @loggerWork('fb')
    def abnormalInfo(self, str_info):
        info = time.strftime('%H:%M:%S ->') + self.name_class + '|异常 —>\n' + str_info
        return info

    @loggerWork('fb')
    def warringInfo(self, str_info):
        info = time.strftime('%H:%M:%S ->') + self.name_class + '|错误 —>\n' + str_info
        return info

    @loggerWork('fb')
    def tranceInfo(self, str_info):
        info = time.strftime('%H:%M:%S ->') + self.name_class + '|回溯 —>\n' + str(*sys.exc_info())
        return info

    @loggerWork('fv')
    def varInfo(self, str_info):
        info = time.strftime('%H:%M:%S ->') + self.name_class + '|变量变更 —>' + str_info
        return info


# 测试函数
def test():
    path_exe = PotsLogFolder.path_exe
    global_tlf = PotsLogFolder.topFolder_log
    global_lf = PotsLogFolder.folder_log
    global_fb = PotsLogFile.flie_bessineseLog
    global_fv = PotsLogFile.file_varLog
    logger_test = PotsLogger('test_test')
    logger_test.getFuncName('test_func')


# 开启动作
if __name__ == '__main__':
    test()
