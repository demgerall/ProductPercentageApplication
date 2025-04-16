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

        """Создание переменных окружения"""
        self.base_save_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        self.search_file_path_Excel = ''

        self.standardSavePathInput.setPlaceholderText(self.base_save_path)

        """Загрузка конфигов"""
        self.parse_config = self.loadParseConfig()
        self.app_config = self.loadAppConfig()

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

    """Загрузка конфига приложения"""
    def loadAppConfig(self) -> object:
        try:
            with open('appConfig.json', 'r') as f:
                app_config = json.load(f)

            if app_config['savePath']:
                self.standardSavePathInput.setPlaceholderText(app_config['savePath'])

            self.fastExportCheckBox.setChecked(app_config['fastExport'] == 'True')
            self.timeDelaySpinBox.setValue(app_config['timeDelay'])

            self.statusLabel.setText('--Загрузка конфига приложения прошла успешно--')

            return app_config

        except Exception as _ex:
            QMessageBox.critical(
                self,
                'Ошибка',
                'Возникла ошибка при загрузке конфига приложения'
            )

            logging.exception(_ex)

        finally:
            f.close()

    """Сохранение конфига приложения"""
    def saveAppConfig(self) -> None:
        is_changed = False

        if self.standardSavePathInput.text():
            is_changed = True

            self.app_config['savePath'] = self.standardSavePathInput.text()
            self.standardSavePathInput.setPlaceholderText(self.standardSavePathInput.text())
            self.standardSavePathInput.clear()

        if str(self.fastExportCheckBox.isChecked()) != self.app_config['fastExport']:
            is_changed = True

            self.app_config['fastExport'] = str(self.fastExportCheckBox.isChecked())

        if self.timeDelaySpinBox.value() != self.app_config['timeDelay']:
            is_changed = True

            self.app_config['timeDelay'] = self.timeDelaySpinBox.value()

        if not is_changed:
            return

        try:
            with open('appConfig.json', 'w') as f:
                json.dump(self.app_config, f)

        except Exception as _ex:
            QMessageBox.critical(
                self,
                'Ошибка',
                'Возникла ошибка при сохранении конфига приложения'
            )

            logging.exception(_ex)

        finally:
            f.close()

    """Загрузка конфига парсера"""
    def loadParseConfig(self) -> object:
        try:
            with open('parserConfig.json', 'r') as f:
                parse_config = json.load(f)

            self.deliveryDateCheckBox.setChecked(parse_config['isDeliveryDateLimit'] == 'True')
            self.deliveryDateSpinBox.setValue(parse_config['deliveryDateLimit'])
            self.instockCheckBox.setChecked(parse_config['onlyInStock'] == 'True')
            self.guaranteeCheckBox.setChecked(parse_config['onlyWithGuarantee'] == 'True')
            self.rateCheckBox.setChecked(parse_config['isStoreRatingLimit'] == 'True')
            self.rateSpinBox.setValue(parse_config['storeRatingLimit'])
            self.blackListCheckBox.setChecked(parse_config['useBlackList'] == 'True')
            self.whiteListCheckBox.setChecked(parse_config['useWhiteList'] == 'True')

            self.statusLabel.setText('--Загрузка конфига парсера прошла успешно--')

            return parse_config

        except Exception as _ex:
            QMessageBox.critical(
                self,
                'Ошибка',
                'Возникла ошибка при загрузке конфига парсера'
            )

            logging.exception(_ex)

        finally:
            f.close()

    """Сохранение конфига парсера"""
    def saveParseConfig(self) -> None:
        is_changed = False

        if str(self.deliveryDateCheckBox.isChecked()) != self.parse_config['isDeliveryDateLimit']:
            is_changed = True

            self.parse_config['isDeliveryDateLimit'] = str(self.deliveryDateCheckBox.isChecked())

        if self.deliveryDateSpinBox != self.parse_config['deliveryDateLimit']:
            is_changed = True

            self.parse_config['deliveryDateLimit'] = self.deliveryDateSpinBox.value()

        if str(self.instockCheckBox.isChecked()) != self.parse_config['onlyInStock']:
            is_changed = True

            self.parse_config['onlyInStock'] = str(self.instockCheckBox.isChecked())

        if str(self.guaranteeCheckBox.isChecked()) != self.parse_config['onlyWithGuarantee']:
            is_changed = True

            self.parse_config['onlyWithGuarantee'] = str(self.guaranteeCheckBox.isChecked())

        if str(self.rateCheckBox.isChecked()) != self.parse_config['isStoreRatingLimit']:
            is_changed = True

            self.parse_config['isStoreRatingLimit'] = str(self.rateCheckBox.isChecked())

        if self.rateSpinBox != self.parse_config['storeRatingLimit']:
            is_changed = True

            self.parse_config['storeRatingLimit'] = self.rateSpinBox.value()

        if str(self.blackListCheckBox.isChecked()) != self.parse_config['useBlackList']:
            is_changed = True

            self.parse_config['useBlackList'] = str(self.blackListCheckBox.isChecked())

        if str(self.whiteListCheckBox.isChecked()) != self.parse_config['useWhiteList']:
            is_changed = True

            self.parse_config['useWhiteList'] = str(self.whiteListCheckBox.isChecked())

        if not is_changed:
            return

        try:
            with open('parserConfig.json', 'w') as f:
                json.dump(self.parse_config, f)

        except Exception as _ex:
            QMessageBox.critical(
                self,
                'Ошибка',
                'Возникла ошибка при сохранении конфига парсера'
            )

            logging.exception(_ex)

        finally:
            f.close()

    def resetParseConfig(self):
        self.search_file_path_Excel = ''
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

    """Сброс пути сохранения на стандартный"""
    def resetStandardSavePath(self):
        self.app_config['savePath'] = ''

        self.standardSavePathInput.setPlaceholderText(self.base_save_path)

    """Загрузка Excel-файла с артикулами"""
    def loadSearchFileExcel(self) -> None:
        filePath, _ = QFileDialog.getOpenFileName(
            self, 'Выберите файл Excel', '', 'Excel Files (*.xlsx)'
        )

        if not filePath:
            self.choosedFileLabel.setText('Файл не выбран')
            return

        self.search_file_path_Excel = filePath

        self.choosedFileLabel.setText(filePath.split('/')[-1])

    """Загрузка Excel-файла для Черного или Белого списка"""

    def importListFileExcel(self, table: QTableWidget) -> None:
        filePath, _ = QFileDialog.getOpenFileName(
            self, 'Выберите файл Excel', '', 'Excel Files (*.xlsx)'
        )

        if not filePath:
            return

        try:
            # Читаем Excel файл
            df = pd.read_excel(filePath)

            # Проверяем, есть ли данные в таблице
            if df.empty:
                QMessageBox.warning(
                    self,
                    'Пустая таблица',
                    'Импортируемый файл не содержит данных.'
                )
                return

            # Проверяем структуру файла
            if len(df.columns) != 2 or list(df.columns) != ['Бренд', 'Магазин']:
                QMessageBox.warning(
                    self,
                    'Ошибка формата',
                    'Импортируемый файл должен содержать 2 колонки с заголовками "Бренд" и "Магазин"'
                )
                return

            # Проверяем наличие пустых строк или ячеек
            has_empty = df.isnull().any().any() or any(df.apply(lambda row: row.str.strip().eq('').any(), axis=1))

            if has_empty:
                reply = QMessageBox.question(
                    self,
                    'Пустые значения',
                    'В таблице обнаружены пустые ячейки или строки. Продолжить импорт? (Пустые строки будут удалены)',
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No
                )

                if reply == QMessageBox.StandardButton.No:
                    return

                # Удаляем строки с пустыми значениями
                df = df.dropna(how='any').reset_index(drop=True)
                df = df[~df.apply(lambda row: row.str.strip().eq('').any(), axis=1)].reset_index(drop=True)

                # Проверяем, остались ли данные после удаления пустых строк
                if df.empty:
                    QMessageBox.warning(
                        self,
                        'Нет данных',
                        'После удаления пустых строк в таблице не осталось данных.'
                    )
                    return

            table.setRowCount(len(df))

            # Заполняем таблицу данными
            for row in range(len(df)):
                for col in range(2):
                    value = str(df.iloc[row, col])
                    item = QTableWidgetItem(value)

                    # Подсветка пустых ячеек (теоретически не должно быть, но на всякий случай оставляем)
                    if not value.strip():
                        item.setBackground(QColor(255, 200, 200))

                    table.setItem(row, col, item)

            self.statusLabel.setText('--Таблица успешно импортирована--')

        except Exception as _ex:
            QMessageBox.critical(
                self,
                'Ошибка',
                f'Не удалось загрузить файл: {str(_ex)}'
            )

            self.statusLabel.setText('--Импорт таблицы не прошел успешно--')
            self.statusLabel.setText('Возникла ошибка. Проверьте файл logs.log')

            logging.exception(_ex)

    """Экспорт таблиц в Excel файл"""
    def exportListFileExcel(self, table: QTableWidget) -> None:
        # Проверяем, не пустая ли таблица
        if table.rowCount() == 0 or table.columnCount() == 0:
            QMessageBox.warning(self, 'Ошибка', 'Таблица пустая! Нет данных для экспорта.')
            return

        # Собираем данные из таблицы и проверяем пустые ячейки
        headers = []
        data = []

        # Получаем заголовки
        for col in range(table.columnCount()):
            headers.append(
                table.horizontalHeaderItem(col).text() if table.horizontalHeaderItem(col) else f'Column {col + 1}')

        # Анализируем данные и добавляем в данные только полностью заполненные строчки
        for row in range(table.rowCount()):
            rowData = []

            for col in range(table.columnCount()):
                item = table.item(row, col)

                if item.text().strip():
                    rowData.append(item.text().strip())
                else:
                    continue

            if len(rowData) == 2:
                data.append(rowData)

        # Проверяем, остались ли данные после удаления
        if not data:
            QMessageBox.warning(
                self,
                'Нет данных для экспорта',
                'После удаления строк с пустыми ячейками таблица стала пустой. Экспорт отменен.'
            )
            return

        if len(data) != table.rowCount():
            reply = QMessageBox.question(
                self,
                'Пустые ячейки обнаружены',
                f'Найдено {table.rowCount() - len(data)} строк с пустыми ячейками. Они будут удалены.\nПродолжить '
                f'экспорт?',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
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
            # Создаем DataFrame и экспортируем
            df = pd.DataFrame(data, columns=headers)
            df.to_excel(filePath, index=False)

            QMessageBox.information(
                self,
                'Экспорт завершен',
                f'Данные успешно экспортированы в Excel!'
            )

        except Exception as _ex:
            QMessageBox.critical(
                self,
                'Ошибка экспорта',
                f'Не удалось экспортировать данные:\n{str(_ex)}'
            )
            logging.exception(_ex)

    """Переход на другую страницу и сохранение данных"""
    def changePage(self, index: int) -> None:
        if self.stackedWidget.currentIndex() == 4:
            self.saveAppConfig()

        if self.stackedWidget.currentIndex() == 1:
            self.saveParseConfig()

        if self.stackedWidget.currentIndex() == 2 or self.stackedWidget.currentIndex() == 3:
            self.updateTableLabels(self.stackedWidget.currentIndex())

        self.stackedWidget.setCurrentIndex(index)

    def updateTableLabels(self, index):
        if index == 2:
            self.blackListEntitiesAmountLabel.setText(f'({self.blackListTable.rowCount()} записи (-ей))')

        if index == 3:
            self.whiteListEntitiesAmountLabel.setText(f'({self.whiteListTable.rowCount()} записи (-ей))')

    """Добавление строки в таблицу"""
    def addTableRow(self, table: QTableWidget) -> None:
        row = table.rowCount()
        table.insertRow(row)
        table.setItem(row, 0, QTableWidgetItem(""))
        table.setItem(row, 1, QTableWidgetItem(""))

    """Удаление выбранных строк из таблицы"""
    def removeTableRow(self, table: QTableWidget) -> None:
        selected_rows = sorted(set(item.row() for item in table.selectedItems()), reverse=True)

        if not selected_rows:
            QMessageBox.warning(self, 'Ошибка', 'Выберите строки для удаления!')
            return

        reply = QMessageBox.question(
            self,
            'Подтверждение',
            f'Удалить выбранные строки ({len(selected_rows)})?',
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
