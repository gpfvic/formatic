import sys,os

from PyQt5 import QtWidgets,QtCore
from PyQt5.QtWidgets import QMainWindow,QApplication

from Main_gui import Ui_MainWindow





if __name__ == "__main__":
    QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    app = QApplication(sys.argv)# 创建QApplication对象，作为GUI主程序入口
    main_gui = Ui_MainWindow()# 创建主窗体对象，实例化Ui_MainWindow    
    qmw = QMainWindow()# 实例化QMainWindow类
    main_gui.setupUi(qmw)# 主窗体对象调用setupUi方法，对QMainWindow对象进行设置
    qmw.show()
    sys.exit(app.exec_())

