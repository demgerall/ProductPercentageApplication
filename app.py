# Инициализация дизайна
# pyuic6 C:/Users/demge/PycharmProjects/ProductPercentageApplication/ui/ProductPercentageApplication.ui -o C:/Users/demge/PycharmProjects/ProductPercentageApplication/ui/ProductPercentageApplicationDesign.py

# Компилятор exe
# pyinstaller -F -w -i "C:/Users/demge/PycharmProjects/ProductPercentageApplication/assets/icons/franz.ico" app.py

import datetime
import time
import os
import sys
import logging

import requests
import pandas as pd

from threading import Thread

from typing import Optional, Any

from dotenv import load_dotenv

from PyQt6 import QtWidgets
from PyQt6.QtCore import QMetaObject, Qt, Q_ARG
from PyQt6.QtWidgets import QFileDialog, QMessageBox, QTableWidgetItem, QTableWidget

from ui import ProductPercentageApplicationDesign

from tools.constants import AppConstants
from tools.tableControl import addTableRow
from tools.dataConvert import tableFromDataframe
from tools.XMLToDict import parseXMLResponseToDict

from configs.configControl import loadParserConfig, loadAppConfig, saveParserConfig, saveAppConfig


def generateColumns(amount: int) -> list[str]:
    """Генерирует список названий колонок для таблицы результатов парсинга.

    Создает стандартный набор колонок для отображения результатов сравнения цен,
    дополняя его динамическими колонками для каждого магазина в указанном количестве.
    Базовая структура колонок берется из AppConstants.COLUMNS['RESULT'].

    Args:
        amount (int): Количество магазинов для которых нужно добавить колонки.
            Должно быть положительным числом (>= 1).

    Returns:
        list[str]: Список названий колонок в формате:
            - Стандартные колонки (из AppConstants)
            - Набор колонок для каждого магазина (7 колонок на магазин):
                * Цена магазина N
                * Кол-во магазина N
                * Описание кол-ва магазина N
                * Название детали магазина N
                * Название магазина N
                * Условия оплаты магазина N
                * Кол-во дней доставки магазина N

    Raises:
        ValueError: Если amount меньше 1

    Examples:
        >>> generateColumns(1)
        [
            'Бренд', 'Артикул', 'Мин НАЛИЧИЕ', 'Сред НАЛИЧИЕ', 'Макс НАЛИЧИЕ',
            'Мин ПОД ЗАКАЗ', 'Сред ПОД ЗАКАЗ', 'Макс ПОД ЗАКАЗ',
            'Цена магазина 1', 'Кол-во магазина 1', 'Описание кол-ва магазина 1',
            'Название детали магазина 1', 'Название магазина 1',
            'Условия оплаты магазина 1', 'Кол-во дней доставки магазина 1'
        ]

        >>> generateColumns(2)[-7:]  # Последние 7 колонок для второго магазина
        [
            'Цена магазина 2', 'Кол-во магазина 2', 'Описание кол-ва магазина 2',
            'Название детали магазина 2', 'Название магазина 2',
            'Условия оплаты магазина 2', 'Кол-во дней доставки магазина 2'
        ]
    """
    if amount < 1:
        raise ValueError(f'Количество магазинов должно быть >= 1, получено {amount}')

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


def createResultsRow(result_data_row: list[str], table: list[dict[str, Any]]) -> list[str]:
    """Формирует строку данных для DataFrame на основе результатов парсинга.

    Расширяет базовую строку результата (result_data_row) данными о товарах из таблицы,
    добавляя для каждой записи из таблицы 7 стандартных полей в строго определенном порядке.

    Args:
        result_data_row (list[str]): Базовая строка с результатами, содержащая:
            - Бренд
            - Артикул
            - Мин/Сред/Макс наличие
            - Мин/Сред/Макс под заказ
        table (list[dict[str, Any]]): Список словарей с данными о товарах, где каждый словарь
            должен содержать следующие ключи:
            - priceV2: Цена товара
            - qty: Количество товара
            - descr_qtyV2: Описание количества
            - class_cat: Категория товара
            - class_user: Название магазина
            - descr_price: Условия оплаты
            - delivery_days: Сроки доставки

    Returns:
        list[str]: Результирующая строка, содержащая:
            - Исходные данные из result_data_row
            - Добавленные данные из table (по 7 полей на каждый элемент)

    Raises:
        KeyError: Если в словаре table отсутствуют обязательные ключи
        TypeError: Если входные параметры не соответствуют ожидаемым типам

    Examples:
        >>> base_row = ["Brand", "Art123", "10", "15", "20", "5", "8", "12"]
        >>> table_data = [
        ...     {
        ...         'priceV2': '100',
        ...         'qty': '5',
        ...         'descr_qtyV2': 'В наличии',
        ...         'class_cat': 'Категория',
        ...         'class_user': 'Магазин1',
        ...         'descr_price': 'Оплата картой',
        ...         'delivery_days': '2'
        ...     }
        ... ]
        >>> createResultsRow(base_row, table_data)
        [
            'Brand', 'Art123', '10', '15', '20', '5', '8', '12',
            '100', '5', 'В наличии', 'Категория', 'Магазин1', 'Оплата картой', '2'
        ]
    """
    if not isinstance(result_data_row, list):
        raise TypeError(f'result_data_row должен быть list, получен {type(result_data_row).__name__}')
    if not isinstance(table, list):
        raise TypeError(f'table должен быть list, получен {type(table).__name__}')

    required_keys = {
        'priceV2', 'qtyV2', 'descr_qtyV2',
        'class_cat', 'class_user',
        'descr_price', 'delivery_days'
    }

    for i, row in enumerate(table, 1):
        if not isinstance(row, dict):
            raise TypeError(f'Элемент {i} в table должен быть dict, получен {type(row).__name__}')
        missing_keys = required_keys - row.keys()
        if missing_keys:
            raise KeyError(f'Отсутствуют обязательные ключи в записи {i}: {missing_keys}')

        result_data_row.extend([
            str(row['priceV2']),
            str(row['qtyV2'] if row['qtyV2'] >= 0 else 0),
            str(row['descr_qtyV2']),
            str(row['class_cat']),
            str(row['class_user']),
            str(row['descr_price']),
            str(row['delivery_days'])
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
        self.parser_config = loadParserConfig(self)
        self.app_config = loadAppConfig(self)

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

    def resetParseConfig(self) -> None:
        """Сбрасывает настройки парсера к значениям по умолчанию и сохраняет конфиг.

        Возвращает все параметры парсера в исходное состояние:
        - Очищает путь к файлу Excel
        - Сбрасывает чекбоксы (доставка, наличие, гарантия, рейтинг)
        - Устанавливает значения спинбоксов по умолчанию
        - Отключает черный/белый списки
        - Сохраняет изменения в конфигурационный файл

        Side effects:
            - Сбрасывает search_file_path_Excel
            - Обновляет текст choosedFileLabel
            - Изменяет состояние всех связанных виджетов
            - Сохраняет изменения через saveParserConfig()
            - Обновляет statusLabel
        """
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

        saveParserConfig(self)
        self.statusLabel.setText('Настройки парсера сброшены к значениям по умолчанию')

    def resetStandardSavePath(self) -> None:
        """Сбрасывает путь сохранения файлов на стандартное значение.

        Устанавливает:
        - Путь сохранения в конфиге приложения в пустую строку
        - Placeholder поля ввода на базовый путь (рабочий стол пользователя)

        Note:
            Не сохраняет конфиг автоматически - требуется вызов saveAppConfig()
        """
        self.app_config['savePath'] = ''
        self.standardSavePathInput.setPlaceholderText(self.base_save_path)
        self.statusLabel.setText('Путь сохранения сброшен на рабочий стол')

    def loadSearchExcelFile(self) -> None:
        """Загружает Excel-файл с артикулами для последующего парсинга.

        Открывает диалоговое окно выбора файла и обрабатывает выбранный файл:
        1. Позволяет пользователю выбрать файл Excel (.xlsx)
        2. Проверяет, что файл был выбран
        3. Сохраняет путь к файлу в search_file_path_Excel
        4. Обновляет интерфейс, отображая имя выбранного файла

        Side effects:
            - Устанавливает значение search_file_path_Excel
            - Обновляет текст choosedFileLabel
            - Может изменить состояние других элементов UI

        Raises:
            ValueError: Если выбранный файл имеет недопустимое расширение

        Examples:
            >>> # Пользователь выбирает файл "products.xlsx"
            >>> self.loadSearchExcelFile()
            >>> print(self.search_file_path_Excel)
            "C:/path/to/products.xlsx"
            >>> print(self.choosedFileLabel.text())
            "products.xlsx"
        """
        file_path, _ = QFileDialog.getOpenFileName(
            parent=self,
            caption='Выберите файл с артикулами',
            directory='',
            filter='Excel Files (*.xlsx);;All Files (*)'
        )

        if not file_path:
            self.choosedFileLabel.setText('Файл не выбран')
            return

        if not file_path.lower().endswith('.xlsx'):
            QMessageBox.warning(
                self,
                'Неверный формат файла',
                'Пожалуйста, выберите файл в формате Excel (.xlsx)'
            )
            return

        self.search_file_path_Excel = file_path
        file_name = os.path.basename(file_path)
        self.choosedFileLabel.setText(file_name)
        self.statusLabel.setText(f'Выбран файл: {file_name}')

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
                    width = min(50, (max_len + 2))
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
                saveParserConfig(self)
            case 1:
                if not self.validateTable(self.brandsTable):
                    return
                saveParserConfig(self)
            case 2:
                if not self.validateTable(self.blackListTable):
                    return
                saveParserConfig(self)
                self.updateTableLabels(self.stackedWidget.currentIndex())
            case 3:
                if not self.validateTable(self.whiteListTable):
                    return
                saveParserConfig(self)
                self.updateTableLabels(self.stackedWidget.currentIndex())
            case 4:
                saveAppConfig(self)

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

        saveParserConfig(self)

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
            logging.error(f'API request failed: {_ex}')
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

            QMetaObject.invokeMethod(
                self.statusLabel,
                'setText',
                Qt.ConnectionType.QueuedConnection,
                Q_ARG(str, f'Артикул {cleaned_article} обрабатывается')
            )

            progress_value = round((i + 1) / len(self.search_file_data) * 100) if round(
                (i + 1) / len(self.search_file_data) * 100) != 100 else 99
            QMetaObject.invokeMethod(
                self.progressBar,
                'setValue',
                Qt.ConnectionType.QueuedConnection,
                Q_ARG(int, progress_value)
            )

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

            response_data_validated = self.validateResult(response_data['table'])

            if not response_data['table'] or not response_data_validated:
                result_data_row.extend(
                    ['Данные отсутствуют']
                )
                result_data_row += [''] * (len(df_success.columns) - len(result_data_row))
                df_success.loc[len(df_success)] = result_data_row

                time.sleep(self.app_config['timeDelay'])
                continue

            if len(response_data_validated) > 10:
                result_data_row = createResultsRow(result_data_row, response_data_validated[:10])
                result_data_row += [''] * (len(df_success.columns) - len(result_data_row))

                df_success.loc[len(df_success)] = result_data_row

                time.sleep(self.app_config['timeDelay'])
                continue

            result_data_row = createResultsRow(result_data_row, response_data_validated)
            result_data_row += [''] * (len(df_success.columns) - len(result_data_row))
            df_success.loc[len(df_success)] = result_data_row

            time.sleep(self.app_config['timeDelay'])

        QMetaObject.invokeMethod(
            self.statusLabel,
            'setText',
            Qt.ConnectionType.QueuedConnection,
            Q_ARG(str, 'Все артикулы обработаны')
        )

        QMetaObject.invokeMethod(
            self.progressBar,
            'setValue',
            Qt.ConnectionType.QueuedConnection,
            Q_ARG(int, 100)
        )

        QMetaObject.invokeMethod(
            self.resultPageButton,
            'setEnabled',
            Qt.ConnectionType.QueuedConnection,
            Q_ARG(bool, True)
        )

        QMetaObject.invokeMethod(
            self.startButton,
            'setEnabled',
            Qt.ConnectionType.QueuedConnection,
            Q_ARG(bool, True)
        )

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
