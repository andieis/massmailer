from PyQt5.QtWidgets import *
from PyQt5.uic import *
from PyQt5.QtCore import Qt


class AboutWindow(QDialog):
    def __init__(self):
        super(AboutWindow, self).__init__()
        loadUi("data/about.ui", self)
        self.buttonBox.rejected.connect(self.exit)
        self.setWindowFlag(Qt.WindowContextHelpButtonHint, False)

    def exit(self):
        self.close()
