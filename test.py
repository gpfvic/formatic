import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QMessageBox
 
class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
 
    def initUI(self):
        self.button = QPushButton('Open Folder', self)
        self.button.clicked.connect(self.openFolder)
        self.button.move(100, 50)
        self.setGeometry(100, 100, 300, 200)
        self.setWindowTitle('Folder Opener')
        self.show()
 
    def openFolder(self):
        directory = QFileDialog.getExistingDirectory(self, "Open Directory", "/home")
        if directory:
            QMessageBox.information(self, "Directory Selected", directory)
 
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())