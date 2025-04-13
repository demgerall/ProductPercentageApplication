# Инициализация дизайна
# pyuic6 C:/Users/demge/PycharmProjects/ProductPercentageApplication/ProductPercentageApplication.ui -o C:/Users/demge/PycharmProjects/ProductPercentageApplication/ProductPercentageApplicationDesign.py

# Компилятор exe
# pyinstaller -F -w -i "C:/Users/demge/PycharmProjects/ProductPercentageApplication/assets/icons/franz.ico" app.py

import datetime
import json
import time
import re
import os
import sys
import logging

from threading import Thread

from PyQt6 import QtWidgets
from PyQt6.QtGui import QFontDatabase, QFont

import ProductPercentageApplicationDesign


class App(QtWidgets.QMainWindow, ProductPercentageApplicationDesign.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        logging.basicConfig(level=logging.DEBUG, filename="logs.log",
                            format="%(levelname)s (%(asctime)s): %(message)s (Line: %(lineno)d) [%(filename)s]",
                            datefmt="%d/%m/%Y %I:%M:%S", encoding='UTF-8', filemode="a")




def main():
    app = QtWidgets.QApplication(sys.argv)
    window = App()
    window.show()
    app.exec()


if __name__ == '__main__':
    main()
