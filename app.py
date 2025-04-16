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


def addTableRow(table: QTableWidget) -> None:
    """Добавление строки в таблицу"""
    row = table.rowCount()

    table.insertRow(row)

    table.setItem(row, 0, QTableWidgetItem(""))
    table.setItem(row, 1, QTableWidgetItem(""))


class App(QtWidgets.QMainWindow, ProductPercentageApplicationDesign.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        """Создание переменных окружения"""
        self.base_save_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        self.search_file_path_Excel = ''

        self.standardSavePathInput.setPlaceholderText(self.base_save_path)

        """Загрузка конфигов"""
        self.parser_config = self.loadParserConfig()
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
        self.chooseFileButton.clicked.connect(self.loadSearchExcelFile)
        self.clearParseSettingsButton.clicked.connect(self.resetParseConfig)

        """Настройка кнопок на странице Замена брендов"""
        self.addTableRowButton.clicked.connect(lambda: addTableRow(self.brandsTable))
        self.deleteTableRowButton.clicked.connect(lambda: self.removeTableRow(self.brandsTable))

        """Настройка кнопок на странице Черный список"""
        self.addBlackListTableRowButton.clicked.connect(lambda: addTableRow(self.blackListTable))
        self.deleteBlackListTableRowButton.clicked.connect(lambda: self.removeTableRow(self.blackListTable))
        self.importBlackListButton.clicked.connect(lambda: self.importListExcelFile(self.blackListTable))
        self.exportBlackListButton.clicked.connect(lambda: self.exportListExcelFile(self.blackListTable))

        """Настройка кнопок на странице Белый список"""
        self.addWhiteListTableRowButton.clicked.connect(lambda: addTableRow(self.whiteListTable))
        self.deleteWhiteListTableRowButton.clicked.connect(lambda: self.removeTableRow(self.whiteListTable))
        self.importWhiteListButton.clicked.connect(lambda: self.importListExcelFile(self.whiteListTable))
        self.exportWhiteListButton.clicked.connect(lambda: self.exportListExcelFile(self.whiteListTable))

    """Загрузка конфига приложения"""
    def loadAppConfig(self) -> object:
        app_config = self.loadConfig('app')

        if app_config['savePath']:
            self.standardSavePathInput.setPlaceholderText(app_config['savePath'])

        self.fastExportCheckBox.setChecked(app_config['fastExport'] == 'True')
        self.timeDelaySpinBox.setValue(app_config['timeDelay'])

        self.statusLabel.setText('--Загрузка конфига приложения прошла успешно--')

        return app_config

    """Загрузка конфига парсера"""
    def loadParserConfig(self) -> object:
        parser_config = self.loadConfig('parser')

        self.deliveryDateCheckBox.setChecked(parser_config['isDeliveryDateLimit'] == 'True')
        self.deliveryDateSpinBox.setValue(parser_config['deliveryDateLimit'])
        self.instockCheckBox.setChecked(parser_config['onlyInStock'] == 'True')
        self.guaranteeCheckBox.setChecked(parser_config['onlyWithGuarantee'] == 'True')
        self.rateCheckBox.setChecked(parser_config['isStoreRatingLimit'] == 'True')
        self.rateSpinBox.setValue(parser_config['storeRatingLimit'])
        self.blackListCheckBox.setChecked(parser_config['useBlackList'] == 'True')
        self.whiteListCheckBox.setChecked(parser_config['useWhiteList'] == 'True')

        self.statusLabel.setText('--Загрузка конфига парсера прошла успешно--')

        return parser_config

    """Загрузка конфигов"""
    def loadConfig(self, config_type: str) -> object:
        config_file_path = 'appConfig.json' if config_type == 'app' else 'parserConfig.json'

        try:
            with open(config_file_path, 'r') as f:
                config = json.load(f)
                return config

        except Exception as _ex:
            QMessageBox.critical(
                self,
                'Ошибка загрузки конфига',
                f'Возникла ошибка при загрузке конфига {'приложения' if config_type == 'app' else 'парсера'}'
            )

            logging.exception(_ex)

        finally:
            f.close()

    """Сохранение конфига приложения"""
    def saveAppConfig(self) -> None:
        is_changed = False

        if self.standardSavePathInput.text().strip():
            is_changed = True

            self.app_config['savePath'] = self.standardSavePathInput.text().strip()
            self.standardSavePathInput.setPlaceholderText(self.standardSavePathInput.text().strip())
            self.standardSavePathInput.clear()

        if str(self.fastExportCheckBox.isChecked()) != self.app_config['fastExport']:
            is_changed = True

            self.app_config['fastExport'] = str(self.fastExportCheckBox.isChecked())

        if self.timeDelaySpinBox.value() != self.app_config['timeDelay']:
            is_changed = True

            self.app_config['timeDelay'] = self.timeDelaySpinBox.value()

        if not is_changed:
            return

        self.saveConfig(self.app_config, 'app')

    """Сохранение конфига парсера"""
    def saveParserConfig(self) -> None:
        is_changed = False

        if str(self.deliveryDateCheckBox.isChecked()) != self.parser_config['isDeliveryDateLimit']:
            is_changed = True

            self.parser_config['isDeliveryDateLimit'] = str(self.deliveryDateCheckBox.isChecked())

        if self.deliveryDateSpinBox != self.parser_config['deliveryDateLimit']:
            is_changed = True

            self.parser_config['deliveryDateLimit'] = self.deliveryDateSpinBox.value()

        if str(self.instockCheckBox.isChecked()) != self.parser_config['onlyInStock']:
            is_changed = True

            self.parser_config['onlyInStock'] = str(self.instockCheckBox.isChecked())

        if str(self.guaranteeCheckBox.isChecked()) != self.parser_config['onlyWithGuarantee']:
            is_changed = True

            self.parser_config['onlyWithGuarantee'] = str(self.guaranteeCheckBox.isChecked())

        if str(self.rateCheckBox.isChecked()) != self.parser_config['isStoreRatingLimit']:
            is_changed = True

            self.parser_config['isStoreRatingLimit'] = str(self.rateCheckBox.isChecked())

        if self.rateSpinBox != self.parser_config['storeRatingLimit']:
            is_changed = True

            self.parser_config['storeRatingLimit'] = self.rateSpinBox.value()

        if str(self.blackListCheckBox.isChecked()) != self.parser_config['useBlackList']:
            is_changed = True

            self.parser_config['useBlackList'] = str(self.blackListCheckBox.isChecked())

        if str(self.whiteListCheckBox.isChecked()) != self.parser_config['useWhiteList']:
            is_changed = True

            self.parser_config['useWhiteList'] = str(self.whiteListCheckBox.isChecked())

        if not is_changed:
            return

        self.saveConfig(self.parser_config, 'parser')

    """Сохранение конфигов"""
    def saveConfig(self, config: object, config_type: str) -> None:
        config_file_path = 'appConfig.json' if config_type == 'app' else 'parserConfig.json'

        try:
            with open(config_file_path, 'w') as f:
                json.dump(config, f)

        except Exception as _ex:
            QMessageBox.critical(
                self,
                'Ошибка сохранения',
                f'Возникла ошибка при сохранении конфига {'приложения' if config_type == 'app' else 'парсера'}'
            )

            logging.exception(_ex)

        finally:
            f.close()

    """Сброс полей настроек парсера и обновление конфига"""
    def resetParseConfig(self) -> None:
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

        self.saveParserConfig()

    """Сброс пути сохранения на стандартный"""
    def resetStandardSavePath(self) -> None:
        self.app_config['savePath'] = ''

        self.standardSavePathInput.setPlaceholderText(self.base_save_path)

    """Загрузка Excel-файла с артикулами"""
    def loadSearchExcelFile(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self, 'Выберите файл Excel', '', 'Excel Files (*.xlsx)'
        )

        if not file_path:
            self.choosedFileLabel.setText('Файл не выбран')
            return

        self.search_file_path_Excel = file_path

        self.choosedFileLabel.setText(file_path.split('/')[-1])

    """Загрузка Excel-файла для Черного или Белого списка"""
    def importListExcelFile(self, table: QTableWidget) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self, 'Выберите файл Excel', '', 'Excel Files (*.xlsx)'
        )

        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)

            if df.empty:
                QMessageBox.warning(
                    self,
                    'Нет данных для импорта',
                    'Импортируемый файл не содержит данных'
                )
                return

            if list(df.columns) != ['Бренд', 'Магазин']:
                QMessageBox.warning(
                    self,
                    'Ошибка формата импортируемой таблицы',
                    'Импортируемый файл должен содержать 2 колонки с заголовками "Бренд" и "Магазин"'
                )
                return

            if df.isnull().any().any() or any(df.apply(lambda row: row.str.strip().eq('').any(), axis=1)):
                df = df.dropna(how='any').reset_index(drop=True)
                df = df[~df.apply(lambda row: row.str.strip().eq('').any(), axis=1)].reset_index(drop=True)

                if df.empty:
                    QMessageBox.warning(
                        self,
                        'Нет данных для импорта',
                        'После удаления пустых строк в таблице не осталось данных'
                    )
                    return

                reply = QMessageBox.question(
                    self,
                    'Обнаружены пустые значения',
                    'В таблице обнаружены пустые ячейки или строки. Продолжить импорт? (Пустые строки будут удалены)',
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No
                )

                if reply == QMessageBox.StandardButton.No:
                    return

            table.setRowCount(len(df))

            for row in range(len(df)):
                for col in range(2):
                    value = str(df.iloc[row, col])
                    item = QTableWidgetItem(value)

                    table.setItem(row, col, item)

            QMessageBox.information(
                self,
                'Импорт завершен',
                f'Данные успешно импортированы в Excel'
            )

        except Exception as _ex:
            QMessageBox.critical(
                self,
                'Ошибка импорта',
                'Не удалось загрузить файл'
            )

            logging.exception(_ex)

    """Экспорт таблиц в Excel файл"""
    def exportListExcelFile(self, table: QTableWidget) -> None:
        if table.rowCount() == 0:
            QMessageBox.warning(self, 'Нет данных для экспорта', 'Экспортируемая таблица не содержит данных')
            return

        headers = ['Бренд', 'Магазин']
        data = []

        for row in range(table.rowCount()):
            row_data = []

            for col in range(table.columnCount()):
                item = table.item(row, col)

                if not item.text().strip():
                    continue

                row_data.append(item.text().strip())

            if len(row_data) == 2:
                data.append(row_data)

        if not data:
            QMessageBox.warning(
                self,
                'Нет данных для экспорта',
                'После удаления строк с пустыми ячейками таблица стала пустой. Экспорт отменен'
            )
            return

        if len(data) != table.rowCount():
            reply = QMessageBox.question(
                self,
                'Обнаружены пустые значения',
                f'В таблице обнаружено {table.rowCount() - len(data)} строк с пустыми значениями. Продолжить импорт? '
                f'(Пустые строки будут удалены)',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )

            if reply == QMessageBox.StandardButton.No:
                return

        # Получаем путь для сохранения файла
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            'Сохранить как Excel',
            '',
            'Excel Files (*.xlsx)'
        )

        if not file_path:
            return

        try:
            # Создаем DataFrame и экспортируем
            df = pd.DataFrame(data, columns=headers)
            df.to_excel(file_path, index=False)

            QMessageBox.information(
                self,
                'Экспорт завершен',
                'Данные успешно экспортированы в Excel'
            )

        except Exception as _ex:
            QMessageBox.critical(
                self,
                'Ошибка экспорта',
                'Не удалось экспортировать данные'
            )
            logging.exception(_ex)

    """Переход на другую страницу и сохранение данных"""
    def changePage(self, index: int) -> None:
        if self.stackedWidget.currentIndex() == 4:
            self.saveAppConfig()

        if self.stackedWidget.currentIndex() == 1:
            self.saveParserConfig()

        if self.stackedWidget.currentIndex() == 2 or self.stackedWidget.currentIndex() == 3:
            self.updateTableLabels(self.stackedWidget.currentIndex())

        self.stackedWidget.setCurrentIndex(index)

    """Обновление лейблов листов в настройках парсера"""
    def updateTableLabels(self, index: int) -> None:
        if index == 2:
            self.blackListEntitiesAmountLabel.setText(f'({self.blackListTable.rowCount()} записи (-ей))')

        if index == 3:
            self.whiteListEntitiesAmountLabel.setText(f'({self.whiteListTable.rowCount()} записи (-ей))')

    """Удаление выбранных строк из таблицы"""
    def removeTableRow(self, table: QTableWidget) -> None:
        selected_rows = sorted(set(item.row() for item in table.selectedItems()), reverse=True)

        if not selected_rows:
            QMessageBox.warning(self, 'Ошибка', 'Выберите строки для удаления')
            return

        reply = QMessageBox.question(
            self,
            'Подтверждение',
            f'Удалить выбранные строки ({len(selected_rows)})?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.No:
            return

        for row in selected_rows:
            table.removeRow(row)


def main() -> None:
    app = QtWidgets.QApplication(sys.argv)

    window = App()
    window.show()

    app.exec()


if __name__ == '__main__':
    main()
