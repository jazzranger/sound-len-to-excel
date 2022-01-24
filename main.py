from os import listdir
from os.path import isfile, join
import sys
from datetime import timedelta

from PyQt5.QtWidgets import QApplication, QFileDialog, QPushButton, QToolTip, QMainWindow
from PyQt5.QtGui import QFont
from mutagen.mp3 import MP3
import xlsxwriter


class App(QMainWindow):

    def __init__(self):
        super().__init__()
        self.title = 'Sound lenght'
        self.left = 300
        self.top = 300
        self.width = 500
        self.height = 200
        self.folderpath = ''
        self.initUI()
    
    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.statusBar()

        QToolTip.setFont(QFont('SansSerif', 10))

        btn = QPushButton('Выбрать папку', self)
        btn.resize(250, 40)
        btn.move(50, 50)
        btn.clicked.connect(self.openFileNameDialog)

        btn = QPushButton('Создать документ', self)
        btn.resize(250, 40)
        btn.move(50, 100)
        btn.clicked.connect(self.create_xlsx)
        
        self.show()

    def openFileNameDialog(self):
        folderpath = QFileDialog.getExistingDirectory(self, 'Select Folder')
        if folderpath:
            self.folderpath = folderpath

        self.statusBar().showMessage(folderpath)

    def create_xlsx(self):
        sound_files = []
        for f in listdir(self.folderpath):
            full_path = join(self.folderpath, f)
            if isfile(full_path) and f.endswith('.mp3'):
                sound_files.append(full_path)

        sound_files.sort()

        sound_len = []
        for path in sound_files:
            audio = MP3(path).info.length
            sound_len.append(str(timedelta(seconds=int(audio))))

        workbook = xlsxwriter.Workbook(f'{self.folderpath}/sound_length.xlsx')
        worksheet = workbook.add_worksheet()

        for i, item in enumerate(sound_len):
            worksheet.write(i, 0, item)

        workbook.close()

        self.statusBar().showMessage("Документ создан")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
