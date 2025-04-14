# Инициализация дизайна
# pyuic6 C:/Users/demge/PycharmProjects/ProductPercentageApplication/ui/ProductPercentageApplication.ui -o C:/Users/demge/PycharmProjects/ProductPercentageApplication/ui/ProductPercentageApplicationDesign.py

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

import pandas as pd

from PyQt6 import QtWidgets
from PyQt6.QtGui import QColor
from PyQt6.QtWidgets import QFileDialog, QMessageBox, QTableWidgetItem, QTableWidget

from ui import ProductPercentageApplicationDesign


class App(QtWidgets.QMainWindow, ProductPercentageApplicationDesign.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        """Создание базовых переменных и загрузка конфигов"""
        self.base_save_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        self.searchFilePathExcel = ''
        self.parseConfig = self.loadParseConfig()
        self.config = self.loadAppConfig()

        """Настройка логирования"""
        logging.basicConfig(level=logging.DEBUG, filename='logs.log',
                            format='%(levelname)s (%(asctime)s): %(message)s (Line: %(lineno)d) [%(filename)s]',
                            datefmt='%d/%m/%Y %I:%M:%S', encoding='UTF-8', filemode='a')

        """Настройка кнопок перехода на страницы в боковом меню"""
        self.parserPageButton.clicked.connect(lambda: self.changePage(0))
        self.brandsPageButton.clicked.connect(lambda: self.changePage(1))
        self.blackListPageButton.clicked.connect(lambda: self.changePage(2))
        self.whiteListPageButton.clicked.connect(lambda: self.changePage(3))
        self.settingsPageButton.clicked.connect(lambda: self.changePage(4))
        self.resultPageButton.clicked.connect(lambda: self.changePage(5))

        """Настройка кнопок на странице Парсинг"""
        self.chooseFileButton.clicked.connect(self.loadSearchFileExcel)

        self.clearParseSettingsButton.clicked.connect(self.resetParseConfig)

        """Настройка кнопок на странице Замена брендов"""
        self.addTableRowButton.clicked.connect(lambda: self.addTableRow(self.brandsTable))
        self.deleteTableRowButton.clicked.connect(lambda: self.removeTableRow(self.brandsTable))

        """Настройка кнопок на странице Черный список"""
        self.addBlackListTableRowButton.clicked.connect(lambda: self.addTableRow(self.blackListTable))
        self.deleteBlackListTableRowButton.clicked.connect(lambda: self.removeTableRow(self.blackListTable))
        self.importBlackListButton.clicked.connect(lambda: self.importListFileExcel(self.blackListTable))
        self.exportBlackListButton.clicked.connect(lambda: self.exportListFileExcel(self.blackListTable))

        """Настройка кнопок на странице Белый список"""
        self.addWhiteListTableRowButton.clicked.connect(lambda: self.addTableRow(self.whiteListTable))
        self.deleteWhiteListTableRowButton.clicked.connect(lambda: self.removeTableRow(self.whiteListTable))
        self.importWhiteListButton.clicked.connect(lambda: self.importListFileExcel(self.whiteListTable))
        self.exportWhiteListButton.clicked.connect(lambda: self.exportListFileExcel(self.whiteListTable))

    def loadAppConfig(self) -> object:
        try:
            with open('appConfig.json', 'r') as f:
                config = json.load(f)

            self.standartSavePathInput.setPlaceholderText(config['savePath'])
            self.fastExportCheckBox.setChecked(config['fastExport'] == 'True')

            self.statusLabel.setText('--Загрузка конфига приложения прошла успешно--')

            return config

        except Exception as _ex:
            self.statusLabel.setText('--Загрузка конфига приложения не прошла успешно--')
            self.statusLabel.setText('Возникла ошибка. Проверьте файл logs.log')

            logging.exception(_ex)

        finally:
            f.close()

    def saveAppConfig(self) -> None:
        isChanged = False

        if self.standartSavePathInput.text():
            isChanged = True

            self.config['savePath'] = self.standartSavePathInput.text()
            self.standartSavePathInput.setPlaceholderText(self.standartSavePathInput.text())
            self.standartSavePathInput.clear()

        if str(self.fastExportCheckBox.isChecked()) != self.config['fastExport']:
            isChanged = True

            self.config['fastExport'] = str(self.fastExportCheckBox.isChecked())

        if isChanged:
            try:
                with open('appConfig.json', 'w') as f:
                    json.dump(self.config, f)

            except Exception as _ex:
                self.statusLabel.setText('--Загрузка конфига приложения не прошла успешно--')
                self.statusLabel.setText('Возникла ошибка. Проверьте файл logs.log')

                logging.exception(_ex)

            finally:
                f.close()

    def loadParseConfig(self) -> object:
        try:
            with open('parserConfig.json', 'r') as f:
                parseConfig = json.load(f)

            self.deliveryDateCheckBox.setChecked(parseConfig['isDeliveryDateLimit'] == 'True')
            self.deliveryDateSpinBox.setValue(parseConfig['deliveryDateLimit'])
            self.instockCheckBox.setChecked(parseConfig['onlyInStock'] == 'True')
            self.guaranteeCheckBox.setChecked(parseConfig['onlyWithGuarantee'] == 'True')
            self.rateCheckBox.setChecked(parseConfig['isStoreRatingLimit'] == 'True')
            self.rateSpinBox.setValue(parseConfig['storeRatingLimit'])
            self.blackListCheckBox.setChecked(parseConfig['useBlackList'] == 'True')
            self.whiteListCheckBox.setChecked(parseConfig['useWhiteList'] == 'True')

            self.statusLabel.setText('--Загрузка конфига парсера прошла успешно--')

            return parseConfig

        except Exception as _ex:
            self.statusLabel.setText('--Загрузка конфига парсера не прошла успешно--')
            self.statusLabel.setText('Возникла ошибка. Проверьте файл logs.log')

            logging.exception(_ex)

        finally:
            f.close()

    def saveParseConfig(self) -> None:
        isChanged = False

        if str(self.deliveryDateCheckBox.isChecked()) != self.parseConfig['isDeliveryDateLimit']:
            isChanged = True

            self.parseConfig['isDeliveryDateLimit'] = str(self.deliveryDateCheckBox.isChecked())

        if self.deliveryDateSpinBox != self.parseConfig['deliveryDateLimit']:
            isChanged = True

            self.parseConfig['deliveryDateLimit'] = self.deliveryDateSpinBox.value()

        if str(self.instockCheckBox.isChecked()) != self.parseConfig['onlyInStock']:
            isChanged = True

            self.parseConfig['onlyInStock'] = str(self.instockCheckBox.isChecked())

        if str(self.guaranteeCheckBox.isChecked()) != self.parseConfig['onlyWithGuarantee']:
            isChanged = True

            self.parseConfig['onlyWithGuarantee'] = str(self.guaranteeCheckBox.isChecked())

        if str(self.rateCheckBox.isChecked()) != self.parseConfig['isStoreRatingLimit']:
            isChanged = True

            self.parseConfig['isStoreRatingLimit'] = str(self.rateCheckBox.isChecked())

        if self.rateSpinBox != self.parseConfig['storeRatingLimit']:
            isChanged = True

            self.parseConfig['storeRatingLimit'] = self.rateSpinBox.value()

        if str(self.blackListCheckBox.isChecked()) != self.parseConfig['useBlackList']:
            isChanged = True

            self.parseConfig['useBlackList'] = str(self.blackListCheckBox.isChecked())

        if str(self.whiteListCheckBox.isChecked()) != self.parseConfig['useWhiteList']:
            isChanged = True

            self.parseConfig['useWhiteList'] = str(self.whiteListCheckBox.isChecked())

        if isChanged:
            try:
                with open('parserConfig.json', 'w') as f:
                    json.dump(self.parseConfig, f)

            except Exception as _ex:
                self.statusLabel.setText('--Загрузка конфига парсера не прошла успешно--')
                self.statusLabel.setText('Возникла ошибка. Проверьте файл logs.log')

                logging.exception(_ex)

            finally:
                f.close()

    def resetParseConfig(self):
        self.searchFilePathExcel = ''
        self.choosedFileLabel.setText('Файл не выбран')
        self.deliveryDateCheckBox.setChecked(False)
        self.deliveryDateSpinBox.setValue(1)
        self.instockCheckBox.setChecked(False)
        self.guaranteeCheckBox.setChecked(False)
        self.rateCheckBox.setChecked(False)
        self.rateSpinBox.setValue(1)
        self.blackListCheckBox.setChecked(False)
        self.whiteListCheckBox.setChecked(False)

        self.saveParseConfig()

    def loadSearchFileExcel(self) -> None:
        """Загрузка Excel-файла с артикулами"""
        filePath, _ = QFileDialog.getOpenFileName(
            self, 'Выберите файл Excel', '', 'Excel Files (*.xlsx)'
        )

        if filePath:
            self.searchFilePathExcel = filePath

            self.choosedFileLabel.setText(filePath.split('/')[-1])

        else:
            self.choosedFileLabel.setText('Файл не выбран')

    def importListFileExcel(self, table: QTableWidget) -> None:
        """Загрузка Excel-файла для Черного или Белого списка"""
        filePath, _ = QFileDialog.getOpenFileName(
            self, 'Выберите файл Excel', '', 'Excel Files (*.xlsx)'
        )

        if not filePath:
            return

        try:
            # Читаем Excel файл
            df = pd.read_excel(filePath)

            # Проверяем структуру файла
            if len(df.columns) != 2 or list(df.columns) != ['Бренд', 'Магазин']:
                QMessageBox.warning(
                    self,
                    'Ошибка формата',
                    'Файл должен содержать 2 колонки с заголовками "Бренд" и "Магазин"'
                )
                return

            table.setRowCount(len(df))

            # Заполняем таблицу  данными
            for row in range(len(df)):
                for col in range(2):
                    value = str(df.iloc[row, col])
                    item = QTableWidgetItem(value)

                    # Подсветка пустых ячеек
                    if not value.strip():
                        item.setBackground(QColor(255, 200, 200))

                    table.setItem(row, col, item)

        except Exception as _ex:
            QMessageBox.critical(
                self,
                'Ошибка',
                f'Не удалось загрузить файл: {str(_ex)}'
            )

            self.statusLabel.setText('--Импорт таблицы не прошел успешно--')
            self.statusLabel.setText('Возникла ошибка. Проверьте файл logs.log')

            logging.exception(_ex)

    def exportListFileExcel(self, table: QTableWidget) -> None:
        # Проверяем, не пустая ли таблица
        if table.rowCount() == 0 or table.columnCount() == 0:
            QMessageBox.warning(self, 'Ошибка', 'Таблица пустая! Нет данных для экспорта.')
            return

        # Проверяем наличие пустых ячеек
        emptyCells = []
        for row in range(table.rowCount()):
            for col in range(table.columnCount()):
                item = table.item(row, col)
                if item is None or item.text().strip() == '':
                    emptyCells.append(f'Строка {row + 1}, Колонка {col + 1}')

        if emptyCells:
            reply = QMessageBox.question(
                self,
                'Пустые ячейки',
                f'Обнаружены пустые ячейки:\n{', '.join(emptyCells)}\n\nПродолжить экспорт?',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                return

        # Получаем путь для сохранения файла
        filePath, _ = QFileDialog.getSaveFileName(
            self,
            'Сохранить как Excel',
            '',
            'Excel Files (*.xlsx)'
        )

        if not filePath:
            return

        try:
            # Собираем данные из таблицы
            headers = []
            data = []

            # Получаем заголовки
            for col in range(table.columnCount()):
                headers.append(
                    table.horizontalHeaderItem(col).text() if table.horizontalHeaderItem(col) else f'Column {col + 1}')

            # Получаем данные
            for row in range(table.rowCount()):
                rowData = []
                for col in range(table.columnCount()):
                    item = table.item(row, col)
                    rowData.append(item.text() if item else '')
                data.append(rowData)

            # Создаем DataFrame
            df = pd.DataFrame(data, columns=headers)

            # Сохраняем в Excel
            df.to_excel(filePath, index=False)

            QMessageBox.information(self, 'Успех', 'Данные успешно экспортированы в Excel!')

        except Exception as _ex:
            QMessageBox.critical(self, 'Ошибка', f'Не удалось экспортировать данные:\n{str(_ex)}')

            logging.exception(_ex)

    def changePage(self, index: int) -> None:
        """Переход на другую страницу и сохранение данных"""
        if self.stackedWidget.currentIndex() == 4:
            self.saveAppConfig()

        if self.stackedWidget.currentIndex() == 1:
            self.saveParseConfig()

        self.stackedWidget.setCurrentIndex(index)

    def addTableRow(self, table: QTableWidget) -> None:
        """Добавление строки в таблицу"""
        row = table.rowCount()
        table.insertRow(row)
        table.setItem(row, 0, QTableWidgetItem(""))
        table.setItem(row, 1, QTableWidgetItem(""))

    def removeTableRow(self, table: QTableWidget) -> None:
        """Удаление выбранных строк из таблицы"""
        selected_rows = sorted(set(item.row() for item in table.selectedItems()), reverse=True)

        if not selected_rows:
            QMessageBox.warning(self, "Ошибка", "Выберите строки для удаления!")
            return

        reply = QMessageBox.question(
            self,
            "Подтверждение",
            f"Удалить выбранные строки ({len(selected_rows)})?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            for row in selected_rows:
                table.removeRow(row)


def main() -> None:
    app = QtWidgets.QApplication(sys.argv)
    window = App()
    window.show()
    app.exec()


if __name__ == '__main__':
    main()
