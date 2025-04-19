# Инициализация дизайна
# pyuic6 C:/Users/demge/PycharmProjects/ProductPercentageApplication/ui/ProductPercentageApplication.ui -o C:/Users/demge/PycharmProjects/ProductPercentageApplication/ui/ProductPercentageApplicationDesign.py

# Компилятор exe
# pyinstaller -F -w -i "C:/Users/demge/PycharmProjects/ProductPercentageApplication/assets/icons/franz.ico" app.py

import datetime
import json
import time
import os
import sys
import logging

from threading import Thread
from typing import Optional

import requests
import pandas as pd
import xml.etree.ElementTree as ET

from dotenv import load_dotenv

from PyQt6 import QtWidgets
from PyQt6.QtWidgets import QFileDialog, QMessageBox, QTableWidgetItem, QTableWidget

from ui import ProductPercentageApplicationDesign


class AppConstants:
    COLUMNS = {
        'SEARCH': ['Производитель', 'Артикул'],
        'LISTS': ['Бренд', 'Магазин'],
        'RESULT': [
            'Бренд', 'Артикул', 'Мин НАЛИЧИЕ', 'Сред НАЛИЧИЕ',
            'Макс НАЛИЧИЕ', 'Мин ПОД ЗАКАЗ', 'Сред ПОД ЗАКАЗ',
            'Макс ПОД ЗАКАЗ'
        ]
    }
    CONFIG_FILES = {
        'app': 'appConfig.json',
        'parser': 'parserConfig.json'
    }
    API_TIMEOUT = 10


def addTableRow(
        table: QTableWidget,
        columns: int = 2,
) -> None:
    """
    Добавляет новую строку с пустыми ячейками в указанную таблицу.

    Вставляет новую строку в конец таблицы и заполняет указанное количество колонок
    пустыми QTableWidgetItem. По умолчанию создает 2 пустые ячейки.

    Args:
        table (QTableWidget): Целевая таблица для добавления строки. Должна содержать
            как минимум 2 колонки (проверяется автоматически).
        columns (int, optional): Количество колонок для инициализации. Не может превышать
            фактическое количество колонок в таблице. По умолчанию 2.

    Returns:
        None: Функция модифицирует переданную таблицу напрямую.

    Raises:
        ValueError: Если таблица содержит меньше 2 колонок или если значение columns
            превышает количество колонок в таблице.

    Examples:
        >>> # Добавление строки в таблицу с 3 колонками
        >>> table = QTableWidget(0, 3)  # 0 строк, 3 колонки
        >>> addTableRow(table, columns=3)
        >>> table.rowCount()
        1

        >>> # Попытка добавить строку в таблицу с 1 колонкой
        >>> addTableRow(QTableWidget(0, 1))
        ValueError: Таблица должна содержать минимум 2 колонки
    """
    if table.columnCount() < 2:
        raise ValueError("Таблица должна содержать минимум 2 колонки")
    if columns > table.columnCount():
        raise ValueError(
            f"Параметр columns ({columns}) превышает количество колонок в таблице "
            f"({table.columnCount()})"
        )

    row = table.rowCount()
    table.insertRow(row)

    for col in range(columns):
        table.setItem(row, col, QTableWidgetItem(""))


def tableToArray(table: QTableWidget) -> list[list[str]]:
    """Преобразует содержимое QTableWidget в двумерный массив строк.

    Проходит по всем ячейкам таблицы и создает двумерный список, где каждый вложенный список
    представляет строку таблицы, а каждый элемент - содержимое соответствующей ячейки.
    Пустые или несуществующие ячейки преобразуются в пустые строки.

    Args:
        table (QTableWidget): Таблица Qt, из которой извлекаются данные. Должна быть
            инициализирована и содержать хотя бы одну строку и колонку для корректной работы.

    Returns:
        list[list[str]]: Двумерный список строк, где:
            - Внешний список содержит строки таблицы
            - Внутренние списки содержат значения ячеек для каждой строки
            - Пустые/несуществующие ячейки представлены пустой строкой ""

    Raises:
        TypeError: Если передан не QTableWidget объект

    Examples:
        >>> table = QTableWidget(2, 3)  # Создаем таблицу 2x3
        >>> table.setItem(0, 0, QTableWidgetItem("Текст"))
        >>> tableToArray(table)
        [['Текст', '', ''], ['', '', '']]

        >>> # Пустая таблица
        >>> table.setRowCount(0)
        >>> tableToArray(table)
        []
    """
    if not isinstance(table, QTableWidget):
        raise TypeError(f"Ожидается QTableWidget, получен {type(table).__name__}")

    result = []

    for row in range(table.rowCount()):
        row_data = []

        for col in range(table.columnCount()):
            item = table.item(row, col)
            text = item.text().strip() if item is not None else ""
            row_data.append(text)

        result.append(row_data)

    return result


def arrayToTable(data: list[list[str]], table: QTableWidget) -> None:
    """Заполняет QTableWidget данными из двумерного массива строк.

    Полностью перезаписывает содержимое таблицы, устанавливая необходимое количество строк и колонок,
    и заполняет ячейки значениями из массива. None значения преобразуются в пустые строки.

    Args:
        data (list[list[str]]): Входные данные для заполнения таблицы, где:
            - Каждый вложенный список представляет строку таблицы
            - Каждый элемент списка представляет значение ячейки
            - None значения автоматически преобразуются в пустые строки
        table (QTableWidget): Целевая таблица для заполнения. Существующее содержимое
            будет полностью очищено перед заполнением.

    Returns:
        None: Метод модифицирует переданную таблицу напрямую.

    Raises:
        TypeError: Если data не является двумерным списком или table не QTableWidget
        ValueError: Если data содержит строки разной длины (не прямоугольная матрица)

    Examples:
        >>> table = QTableWidget()
        >>> data = [["A1", "B1"], ["A2", "B2"]]
        >>> arrayToTable(data, table)
        >>> table.rowCount()
        2
        >>> table.columnCount()
        2

        >>> # Обработка None значений
        >>> arrayToTable([["X", None], [None, "Y"]], table)
        >>> table.item(0, 1).text()
        ''
    """
    if not isinstance(table, QTableWidget):
        raise TypeError(f"Ожидается QTableWidget, получен {type(table).__name__}")
    if not all(isinstance(row, list) for row in data):
        raise TypeError("Data должен быть двумерным списком")
    if data and len({len(row) for row in data}) > 1:
        raise ValueError("Все строки в data должны иметь одинаковую длину")

    table.setRowCount(0)

    if not data:
        return

    table.setRowCount(len(data))
    table.setColumnCount(len(data[0]))

    for row_idx, row_data in enumerate(data):
        for col_idx, cell_data in enumerate(row_data):
            item = QTableWidgetItem(str(cell_data) if cell_data is not None else "")
            table.setItem(row_idx, col_idx, item)


def tableToDict(table: QTableWidget) -> dict[str, str]:
    """
    Преобразует таблицу в словарь, где:
    - Первый столбец (колонка 0) — ключи словаря
    - Второй столбец (колонка 1) — значения
    """
    result = {}

    for row in range(table.rowCount()):
        key_item = table.item(row, 0)
        key = key_item.text() if key_item is not None else ""

        value_item = table.item(row, 1)
        value = value_item.text() if value_item is not None else ""

        if key:
            result[key] = value

    return result


def dictToTable(data: dict, table: QTableWidget) -> None:
    """
    Заполняет таблицу данными из словаря.
    Ключи помещаются в первый столбец, значения - во второй.
    """
    table.setRowCount(len(data))

    for row, (key, value) in enumerate(data.items()):
        key_item = QTableWidgetItem(str(key))
        table.setItem(row, 0, key_item)

        value_item = QTableWidgetItem(str(value) if value is not None else "")
        table.setItem(row, 1, value_item)


def tableFromDataframe(table: QTableWidget, data: pd.DataFrame) -> None:
    """Заполняет таблицу данными из pandas DataFrame"""
    table.clear()

    n_rows, n_cols = data.shape
    table.setRowCount(n_rows)
    table.setColumnCount(n_cols)

    table.setHorizontalHeaderLabels(data.columns.tolist())

    for i in range(n_rows):
        for j in range(n_cols):
            value = data.iloc[i, j]

            # Преобразуем значение в строку (если не NaN)
            item_value = str(value) if not pd.isna(value) else ""

            # Создаём элемент таблицы
            item = QTableWidgetItem(item_value)

            # Вставляем элемент в таблицу
            table.setItem(i, j, item)

    table.resizeColumnsToContents()


def parseXMLResponseToDict(response: requests.Response) -> dict:
    """Переводит XML в словарь"""
    root = ET.fromstring(response.text)

    json_str = root.text.strip()

    try:
        return json.loads(json_str)
    except json.JSONDecodeError as _ex:
        raise ValueError(f"Не удалось распарсить JSON из XML: {_ex}")


def generateColumns(amount: int) -> list[str]:
    """
    Создаем список названий колонок
    """
    columns = AppConstants.COLUMNS['RESULT']

    for i in range(1, amount + 1):
        columns.extend([
            f'Цена магазина {i}',
            f'Кол-во магазина {i}',
            f'Описание кол-ва магазина {i}',
            f'Название детали магазина {i}',
            f'Название магазина {i}',
            f'Условия оплаты магазина {i}',
            f'Кол-во дней доставки магазина {i}'
        ])

    return columns


def createResultsRow(result_data_row: list[str], table: list[dict]) -> list[str]:
    """
    Создаем строку с результатом для Dataframe
    """
    for table_row in table:
        result_data_row.extend([
            table_row['priceV2'],
            table_row['qty'],
            table_row['descr_qtyV2'],
            table_row['class_cat'],
            table_row['class_user'],
            table_row['descr_price'],
            table_row['delivery_days']
        ])

    return result_data_row


class App(QtWidgets.QMainWindow, ProductPercentageApplicationDesign.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        """Загружаем переменные из .env"""
        load_dotenv()
        self.api_keys = os.getenv('API_KEYS').split(' ')
        self.api_url = os.getenv('API_URL')

        """Создание переменных окружения"""
        self.base_save_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        self.search_file_path_Excel = ''
        self.search_file_data = []
        self.result_data = None

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
        self.startButton.clicked.connect(self.prepareForStartParsing)

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

        """Настройка кнопок на страницу Результаты"""
        self.exportResultsButton.clicked.connect(lambda: self.exportResultExcelFile('standard'))
        self.exportResultsAsButton.clicked.connect(lambda: self.exportResultExcelFile('as'))

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
        dictToTable(parser_config['brandsList'], self.brandsTable)
        arrayToTable(parser_config['blackList'], self.blackListTable)
        arrayToTable(parser_config['whiteList'], self.whiteListTable)

        self.statusLabel.setText('--Загрузка конфига парсера прошла успешно--')

        return parser_config

    """Загрузка конфигов"""

    def loadConfig(self, config_type: str) -> object:
        config_file_path = AppConstants.CONFIG_FILES[config_type]

        try:
            with open(config_file_path, 'r') as f:
                config = json.load(f)
                return config

        except Exception as _ex:
            logging.exception(_ex)
            QMessageBox.critical(
                self,
                'Ошибка загрузки конфига',
                f'Возникла ошибка при загрузке конфига {"приложения" if config_type == "app" else "парсера"}'
            )

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

        self.parser_config['brandsList'] = tableToDict(self.brandsTable)
        self.parser_config['blackList'] = tableToArray(self.blackListTable)
        self.parser_config['whiteList'] = tableToArray(self.whiteListTable)

        if not is_changed:
            return

        self.saveConfig(self.parser_config, 'parser')

    """Сохранение конфигов"""

    def saveConfig(self, config: object, config_type: str) -> None:
        config_file_path = AppConstants.CONFIG_FILES[config_type]

        try:
            with open(config_file_path, 'w') as f:
                json.dump(config, f)

        except Exception as _ex:
            logging.exception(_ex)
            QMessageBox.critical(
                self,
                'Ошибка сохранения',
                f'Возникла ошибка при сохранении конфига {"приложения" if config_type == "app" else "парсера"}'
            )

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

    """Загрузка данных из Excel-файла с артикулами и их валидация"""

    def importSearchExcelFileToArray(self, path: str) -> list[list[str]] or None:
        try:
            df = pd.read_excel(path)

            if df.empty:
                QMessageBox.warning(
                    self,
                    'Нет данных для импорта',
                    'Импортируемый файл не содержит данных'
                )
                return

            if list(df.columns) != AppConstants.COLUMNS['SEARCH']:
                QMessageBox.warning(
                    self,
                    'Ошибка формата импортируемой таблицы',
                    'Импортируемый файл должен содержать 2 колонки с заголовками "Производитель" и "Артикул"'
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

            return df.values.tolist()

        except Exception as _ex:
            logging.exception(_ex)
            QMessageBox.critical(
                self,
                'Ошибка импорта',
                'Не удалось загрузить файл'
            )

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

            if list(df.columns) != AppConstants.COLUMNS['LISTS']:
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
                'Данные успешно импортированы в Excel'
            )

        except Exception as _ex:
            logging.exception(_ex)
            QMessageBox.critical(
                self,
                'Ошибка импорта',
                'Не удалось загрузить файл'
            )

    """Экспорт таблиц в Excel файл c валидацией"""

    def exportListExcelFile(self, table: QTableWidget) -> None:
        if table.rowCount() == 0:
            QMessageBox.warning(self, 'Нет данных для экспорта', 'Экспортируемая таблица не содержит данных')
            return

        headers = AppConstants.COLUMNS['LISTS']
        data = []

        for row in range(table.rowCount()):
            row_data = []

            for col in range(table.columnCount()):
                item = table.item(row, col)

                if not item.text().strip() or item is None:
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

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            'Сохранить как Excel',
            '',
            'Excel Files (*.xlsx)'
        )

        if not file_path:
            return

        try:
            df = pd.DataFrame(data, columns=headers)
            df.to_excel(file_path, index=False)

            QMessageBox.information(
                self,
                'Экспорт завершен',
                'Данные успешно экспортированы в Excel'
            )

        except Exception as _ex:
            logging.exception(_ex)
            QMessageBox.critical(
                self,
                'Ошибка экспорта',
                'Не удалось экспортировать данные'
            )

    """Экспорт таблиц в Excel файл без валидации"""

    def exportResultExcelFile(self, save_type: str) -> None:
        file_name = f'Проценка товара от {datetime.datetime.now().strftime("%d-%b-%Y %H-%M-%S")}.xlsx'
        if save_type == 'standard':
            file_path = f'{self.base_save_path}/{file_name}' if not self.app_config[
                'savePath'] else f'{self.app_config['savePath']}/{file_name}'
        else:
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                'Сохранить как Excel',
                '',
                'Excel Files (*.xlsx)'
            )

            if not file_path:
                return

        try:
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                self.result_data.to_excel(writer, index=False, sheet_name='Prices')

                workbook = writer.book
                worksheet = writer.sheets['Prices']

                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'border': 1
                })

                special_header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'border': 1,
                    'bg_color': '#607ebc',
                    'font_color': '#faf5ee'
                })

                data_format = workbook.add_format({
                    'text_wrap': True,
                    'valign': 'vcenter',
                    'font_size': 10,
                    'border': 1
                })

                for col_num, value in enumerate(self.result_data.columns.values):
                    fmt = special_header_format if value in [
                        'Мин НАЛИЧИЕ', 'Сред НАЛИЧИЕ', 'Макс НАЛИЧИЕ',
                        'Мин ПОД ЗАКАЗ', 'Сред ПОД ЗАКАЗ', 'Макс ПОД ЗАКАЗ'
                    ] else header_format
                    worksheet.write(0, col_num, value, fmt)

                for row in range(1, len(self.result_data) + 1):
                    for col in range(len(self.result_data.columns)):
                        worksheet.write(row, col, self.result_data.iloc[row - 1, col], data_format)

                for i, column in enumerate(self.result_data.columns):
                    max_len = max(self.result_data[column].astype(str).map(len).max(), len(column))
                    width = min(50, (max_len + 2) * 1.2)
                    worksheet.set_column(i, i, width)

                worksheet.freeze_panes(1, 0)

        except Exception as _ex:
            logging.exception(_ex)
            QMessageBox.critical(
                self,
                'Ошибка экспорта',
                'Не удалось экспортировать данные'
            )

    """Переход на другую страницу и сохранение данных"""

    def changePage(self, index: int) -> None:
        match self.stackedWidget.currentIndex():
            case 0:
                self.saveParserConfig()
            case 1:
                if not self.validateTable(self.brandsTable):
                    return
                self.saveParserConfig()
            case 2:
                if not self.validateTable(self.blackListTable):
                    return
                self.saveParserConfig()
                self.updateTableLabels(self.stackedWidget.currentIndex())
            case 3:
                if not self.validateTable(self.whiteListTable):
                    return
                self.saveParserConfig()
                self.updateTableLabels(self.stackedWidget.currentIndex())
            case 4:
                self.saveAppConfig()

        self.stackedWidget.setCurrentIndex(index)

    """Валидация таблиц: удаление строк с пустыми ячейками"""

    def validateTable(self, table: QTableWidget) -> bool:
        if table.rowCount() == 0:
            return True

        rows_to_remove = set()

        for row in range(table.rowCount()):
            has_empty = False

            for col in range(table.columnCount()):
                item = table.item(row, col)

                if item is None or item.text().strip() == '':
                    has_empty = True
                    break

            if has_empty:
                rows_to_remove.add(row)

        if not rows_to_remove:
            return True

        reply = QMessageBox.question(
            self,
            'Подтверждение',
            'В таблице есть пустые значения, после перехода они будут удалены. Продолжить?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.No:
            return False

        for row in sorted(rows_to_remove, reverse=True):
            table.removeRow(row)

        return True

    """Обновление лейблов листов в настройках парсера"""

    def updateTableLabels(self, index: int) -> None:
        if index == 2:
            self.blackListEntitiesAmountLabel.setText(f'({self.blackListTable.rowCount()} записи (-ей))')

        if index == 3:
            self.whiteListEntitiesAmountLabel.setText(f'({self.whiteListTable.rowCount()} записи (-ей))')

    """Удаление выбранных строк из таблицы"""

    def removeTableRow(self, table: QTableWidget) -> None:
        selected_rows = set(item.row() for item in table.selectedItems())

        if not selected_rows:
            QMessageBox.warning(self, 'Ошибка', 'Выберите строки для удаления')
            return

        reply = QMessageBox.question(
            self,
            'Подтверждение',
            'Удалить выбранные строки?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.No:
            return

        for row in sorted(selected_rows, reverse=True):
            table.removeRow(row)

    """Подготовка данных парсера перед стартом парсинга"""

    def prepareForStartParsing(self):
        self.progressBar.setValue(0)

        with open("logs.log", "w", encoding='UTF-8') as f:
            f.write("")
        f.close()

        if not self.api_keys:
            QMessageBox.critical(self, "Ошибка", "Ключи API не обнаружены")
            return

        if not self.search_file_path_Excel:
            QMessageBox.critical(self, "Ошибка", "Выберите файл Excel!")
            return

        self.search_file_data = self.importSearchExcelFileToArray(self.search_file_path_Excel)
        if not self.search_file_data:
            self.search_file_path_Excel = ''
            self.choosedFileLabel.setText('Файл не выбран')
            return

        self.saveParserConfig()

        self.startButton.setEnabled(False)

        thread = Thread(target=self.startParsing, daemon=True)
        thread.start()

    def validateResult(self, response_data_table: list[dict]) -> list[dict]:
        results = []

        for i in range(len(response_data_table)):
            if self.parser_config['isDeliveryDateLimit'] == 'True':
                if self.parser_config['deliveryDateLimit'] < response_data_table[i]['delivery_days']:
                    continue
            if self.parser_config['onlyInStock'] == 'True':
                if response_data_table[i]['instock'] != 1:
                    continue
            # if self.parser_config['onlyWithGuarantee'] == 'True':
            #
            if self.parser_config['isStoreRatingLimit'] == 'True':
                if self.parser_config['storeRatingLimit'] > response_data_table[i]['rating']:
                    continue
            if self.parser_config['useBlackList'] == 'True' and len(self.parser_config['blackList']) != 0:
                for j in range(len(self.parser_config['blackList'])):
                    if (self.parser_config['blackList'][j][0] == response_data_table[i]['class_man'] and
                            self.parser_config['blackList'][j][1] == response_data_table[i]['class_user']):
                        continue
            if self.parser_config['useWhiteList'] == 'True' and len(self.parser_config['whiteList']) != 0:
                for j in range(len(self.parser_config['whiteList'])):
                    if (self.parser_config['whiteList'][j][0] != response_data_table[i]['class_man'] or
                            self.parser_config['whiteList'][j][1] != response_data_table[i]['class_user']):
                        continue

            results.append(response_data_table[i])

        return results

    def safeAPIRequest(self, params: dict) -> Optional[dict]:
        try:
            response = requests.get(
                self.api_url,
                params=params,
                timeout=AppConstants.API_TIMEOUT,
                verify=True
            )
            response.raise_for_status()
            return parseXMLResponseToDict(response)

        except (requests.RequestException, ValueError) as _ex:
            logging.error(f"API request failed: {_ex}")
            return

    def startParsing(self):
        df_errors = pd.DataFrame(columns=AppConstants.COLUMNS['SEARCH'])
        df_success = pd.DataFrame(columns=generateColumns(10))

        for i in range(len(self.search_file_data)):
            cleaned_article = str(int(self.search_file_data[i][1])).replace('#', '')

            if self.parser_config['brandsList'].get(self.search_file_data[i][0]):
                normalized_brand = self.parser_config['brandsList'].get(self.search_file_data[i][0])

            else:
                normalized_brand = self.search_file_data[i][0]

            # self.statusLabel.setText(f'Артикул {cleaned_article} обрабатывается')
            # self.progressBar.setValue(round((i + 1) / len(self.search_file_data) * 100) if round(
            #     (i + 1) / len(self.search_file_data) * 100) != 100 else 99)

            params = {
                'api_key': self.api_keys[i % len(self.api_keys)],
                'code_region': self.parser_config['regionCode'],
                'partnumber': cleaned_article,
                'class_man': normalized_brand,
                "type_request": 5,
                'login': '',
                'password': '',
                'search_text': cleaned_article,
                'row_count': 500
            }

            response_data = self.safeAPIRequest(params)

            if not response_data:
                df_errors.loc[len(df_errors)] = [normalized_brand, cleaned_article]

                time.sleep(self.app_config['timeDelay'])
                continue

            result_data_row = [
                normalized_brand,
                cleaned_article,
                response_data['price_min_instock'],
                response_data['price_avg_instock'],
                response_data['price_max_instock'],
                response_data['price_min_order'],
                response_data['price_avg_order'],
                response_data['price_max_order'],
            ]

            if not response_data['table']:
                result_data_row.extend(
                    ['Данные отсутствуют']
                )
                result_data_row += [''] * (len(df_success.columns) - len(result_data_row))
                df_success.loc[len(df_success)] = result_data_row

                time.sleep(self.app_config['timeDelay'])
                continue

            response_data_validated = self.validateResult(response_data['table'])

            if len(response_data_validated) > 10:
                result_data_row = createResultsRow(result_data_row, response_data_validated[:10])

                df_success.loc[len(df_success)] = result_data_row

                time.sleep(self.app_config['timeDelay'])
                continue

            result_data_row = createResultsRow(result_data_row, response_data_validated)
            result_data_row += [''] * (len(df_success.columns) - len(result_data_row))
            df_success.loc[len(df_success)] = result_data_row

            time.sleep(self.app_config['timeDelay'])

        # self.statusLabel.setText('Все артикулы обработаны')
        # self.progressBar.setValue(100)

        self.resultPageButton.setEnabled(True)
        self.startButton.setEnabled(True)

        self.result_data = df_success

        tableFromDataframe(self.resultsTable, self.result_data)
        self.stackedWidget.setCurrentIndex(5)

        if self.app_config['fastExport'] == 'True':
            self.exportResultExcelFile('standard')


def main() -> None:
    app = QtWidgets.QApplication(sys.argv)

    window = App()
    window.show()

    app.exec()


if __name__ == '__main__':
    main()
