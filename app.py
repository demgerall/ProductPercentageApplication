# Инициализация дизайна
# pyuic6 C:/Users/demge/PycharmProjects/ProductPercentageApplication/ui/ProductPercentageApplication.ui -o C:/Users/demge/PycharmProjects/ProductPercentageApplication/ui/ProductPercentageApplicationDesign.py

# Компилятор exe
# pyinstaller -F -w -i "C:/Users/demge/PycharmProjects/ProductPercentageApplication/assets/icons/franz.ico" app.py

import time
import os
import sys
import logging

import pandas as pd

from threading import Thread

from dotenv import load_dotenv

from PyQt6 import QtWidgets
from PyQt6.QtCore import QMetaObject, Qt, Q_ARG
from PyQt6.QtWidgets import QMessageBox

from ui import ProductPercentageApplicationDesign

from configs.configControl import loadParserConfig, loadAppConfig, saveParserConfig

from tools.constants import AppConstants
from tools.dataConvert import tableFromDataframe

from tools.appControl import changePage, updateTableLabels
from tools.resetsTools import resetParseConfig, resetStandardSavePath
from tools.tableControl import addTableRow, removeTableRow

from tools.exportControl import exportListExcelFile, exportErrorArticlesExcelFile, exportResultExcelFile
from tools.importControl import importListExcelFile, importSearchExcelFileToArray, loadSearchExcelFilePath

from tools.APIRequst import safeAPIRequest
from tools.resultControl import generateColumns, validateResult, createResultsRow


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

        updateTableLabels(self, 2)
        updateTableLabels(self, 3)

        """Настройка логирования"""
        logging.basicConfig(level=logging.DEBUG, filename='logs.log',
                            format='%(levelname)s (%(asctime)s): %(message)s (Line: %(lineno)d) [%(filename)s]',
                            datefmt='%d/%m/%Y %I:%M:%S', encoding='UTF-8', filemode='a')

        with open('logs.log', 'w', encoding='UTF-8') as log_file:
            log_file.write('')

        """Настройка кнопок перехода на страницы в боковом меню"""
        self.parserPageButton.clicked.connect(lambda: changePage(self, 0))
        self.brandsPageButton.clicked.connect(lambda: changePage(self, 1))
        self.blackListPageButton.clicked.connect(lambda: changePage(self, 2))
        self.whiteListPageButton.clicked.connect(lambda: changePage(self, 3))
        self.settingsPageButton.clicked.connect(lambda: changePage(self, 4))
        self.resultPageButton.clicked.connect(lambda: changePage(self, 5))

        """Настройка кнопок на странице Парсинг"""
        self.chooseFileButton.clicked.connect(lambda: loadSearchExcelFilePath(self))
        self.clearParseSettingsButton.clicked.connect(lambda: resetParseConfig(self))
        self.startButton.clicked.connect(self.prepare)

        """Настройка кнопок на странице Замена брендов"""
        self.addTableRowButton.clicked.connect(lambda: addTableRow(self.brandsTable))
        self.deleteTableRowButton.clicked.connect(lambda: removeTableRow(self, self.brandsTable))

        """Настройка кнопок на странице Черный список"""
        self.addBlackListTableRowButton.clicked.connect(lambda: addTableRow(self.blackListTable))
        self.deleteBlackListTableRowButton.clicked.connect(lambda: removeTableRow(self, self.blackListTable))
        self.importBlackListButton.clicked.connect(lambda: importListExcelFile(self, self.blackListTable))
        self.exportBlackListButton.clicked.connect(lambda: exportListExcelFile(self, self.blackListTable, 'black'))

        """Настройка кнопок на странице Белый список"""
        self.addWhiteListTableRowButton.clicked.connect(lambda: addTableRow(self.whiteListTable))
        self.deleteWhiteListTableRowButton.clicked.connect(lambda: removeTableRow(self, self.whiteListTable))
        self.importWhiteListButton.clicked.connect(lambda: importListExcelFile(self, self.whiteListTable))
        self.exportWhiteListButton.clicked.connect(lambda: exportListExcelFile(self, self.whiteListTable, 'white'))

        """Настройка кнопок на страницу Результаты"""
        self.exportResultsButton.clicked.connect(lambda: exportResultExcelFile(self, 'standard'))
        self.exportResultsAsButton.clicked.connect(lambda: exportResultExcelFile(self, 'as'))

        """Настройка кнопок на странице Настроек"""
        self.clearStandardSavePathButton.clicked.connect(lambda: resetStandardSavePath(self))

    def prepare(self) -> None:
        """
        Подготавливает систему к началу парсинга: выполняет предварительные проверки,
        сбрасывает состояние интерфейса, запускает процесс парсинга в отдельном потоке.

        Checks:
            - Наличие ключей API (self.api_keys)
            - Выбор файла Excel (self.search_file_path_Excel)
            - Корректность загруженных данных (self.search_file_data)

        Note:
            - Сбрасывает прогресс-бар (self.progressBar)
            - Блокирует кнопки (self.resultPageButton, self.startButton)
            - Очищает файл логов (logs.log)
            - Сохраняет конфигурацию (saveParserConfig)
            - Запускает парсинг в потоке (self.startParsing)
        """
        self.progressBar.setValue(0)
        self.resultPageButton.setEnabled(False)

        with open('logs.log', 'w', encoding='UTF-8') as log_file:
            log_file.write('')

        if not self.api_keys:
            QMessageBox.critical(self, 'Ошибка запуска', 'Необходимо указать ключи API для работы парсера')
            return

        if not self.search_file_path_Excel:
            QMessageBox.critical(self, 'Ошибка файла', 'Не выбран файл Excel с данными для парсинга')
            return

        try:
            self.search_file_data = importSearchExcelFileToArray(self, self.search_file_path_Excel)
            if not self.search_file_data or len(self.search_file_data) == 0:
                self.search_file_path_Excel = ''
                self.choosedFileLabel.setText('Файл не содержит данных')
                QMessageBox.warning(self, 'Ошибка данных', 'Выбранный файл не содержит данных для обработки')
                return
        except Exception as ex:
            self.search_file_path_Excel = ''
            self.choosedFileLabel.setText('Ошибка загрузки')
            QMessageBox.critical(self, 'Ошибка', f'Не удалось загрузить данные из файла: {str(ex)}')
            return

        saveParserConfig(self)

        self.clearParseSettingsButton.setEnabled(False)
        self.startButton.setEnabled(False)

        try:
            thread = Thread(target=self.run, daemon=True)
            thread.start()
        except Exception as ex:
            QMessageBox.critical(self, 'Ошибка', f'Не удалось запустить поток парсинга: {str(ex)}')
            self.startButton.setEnabled(True)

    def run(self) -> None:
        """
        Основной метод парсинга, выполняемый в отдельном потоке.
        Обрабатывает данные из self.search_file_data, выполняет API-запросы,
        сохраняет результаты в self.result_data и обновляет интерфейс.

        Note:
            1. Создает DataFrame для результатов (df_success) и ошибок (df_errors)
            2. Для каждого элемента в search_file_data:
               - Нормализует артикул и бренд
               - Обновляет статус в интерфейсе
               - Формирует и отправляет API-запрос
               - Обрабатывает ответ (success/error)
               - Сохраняет результаты
               - Выдерживает паузу между запросами
            3. По завершении обновляет интерфейс и сохраняет результаты
        """
        try:
            df_errors = pd.DataFrame(columns=AppConstants.COLUMNS['SEARCH'])
            df_success = pd.DataFrame(columns=generateColumns(10))
            total_items = len(self.search_file_data)

            QMetaObject.invokeMethod(
                self.progressBar,
                'setValue',
                Qt.ConnectionType.QueuedConnection,
                Q_ARG(int, 1)
            )

            for i, (brand, article) in enumerate((item[0], str(item[1]).replace('#', '')) for item
                                                 in self.search_file_data):
                try:
                    normalized_brand = self.parser_config['brandsList'].get(brand, brand)

                    QMetaObject.invokeMethod(
                        self.statusLabel,
                        'setText',
                        Qt.ConnectionType.QueuedConnection,
                        Q_ARG(str, f'Артикул {article} обрабатывается ({i + 1} из {total_items})')
                    )

                    progress_value = min(99, round((i + 1) / total_items * 100 + 1))
                    QMetaObject.invokeMethod(
                        self.progressBar,
                        'setValue',
                        Qt.ConnectionType.QueuedConnection,
                        Q_ARG(int, progress_value)
                    )

                    params = {
                        'api_key': self.api_keys[i % len(self.api_keys)],
                        'code_region': self.parser_config['regionCode'],
                        'partnumber': article,
                        'class_man': normalized_brand,
                        "type_request": 5,
                        'login': '',
                        'password': '',
                        'search_text': article,
                        'row_count': 500
                    }

                    response_data = safeAPIRequest(self, params)

                    if not response_data:
                        df_errors.loc[len(df_errors)] = [normalized_brand, article]

                        time.sleep(self.app_config['timeDelay'])
                        continue

                    result_row = [
                        normalized_brand,
                        article,
                        response_data['price_min_instock'],
                        response_data['price_avg_instock'],
                        response_data['price_max_instock'],
                        response_data['price_min_order'],
                        response_data['price_avg_order'],
                        response_data['price_max_order'],
                    ]
                    validated_data = validateResult(self, response_data.get('table', []))

                    if not validated_data:
                        result_row.extend(['Данные отсутствуют'])
                        result_row += [''] * (len(df_success.columns) - len(result_row))
                        df_success.loc[len(df_success)] = result_row

                        time.sleep(self.app_config['timeDelay'])
                        continue

                    result_data = validated_data[:10] if len(validated_data) > 10 else validated_data
                    result_row = createResultsRow(result_row, result_data)

                    if len(validated_data) < 10:
                        result_row.extend(['Больше данных нет'])
                        result_row += [''] * (len(df_success.columns) - len(result_row))

                    df_success.loc[len(df_success)] = result_row

                    time.sleep(self.app_config['timeDelay'])

                except Exception as ex:
                    logging.error(f'Ошибка обработки артикула {article}: {str(ex)}')

                    time.sleep(self.app_config['timeDelay'])
                    continue

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
            QMetaObject.invokeMethod(
                self.clearParseSettingsButton,
                'setEnabled',
                Qt.ConnectionType.QueuedConnection,
                Q_ARG(bool, True)
            )

            self.result_data = df_success
            tableFromDataframe(self.resultsTable, self.result_data)

            self.stackedWidget.setCurrentIndex(5)

            exportErrorArticlesExcelFile(self, df_errors)

            if self.app_config['fastExport'] == 'True':
                exportResultExcelFile(self, 'standard')

        except Exception as ex:
            logging.error(f'Ошибка внутри потока: {str(ex)}')
            QMessageBox.critical(self, 'Ошибка', f'Не удалось запустить поток парсинга: {str(ex)}')

            QMetaObject.invokeMethod(
                self.statusLabel,
                'setText',
                Qt.ConnectionType.QueuedConnection,
                Q_ARG(str, 'Не удалось запустить поток парсинга')
            )
            QMetaObject.invokeMethod(
                self.progressBar,
                'setValue',
                Qt.ConnectionType.QueuedConnection,
                Q_ARG(int, 0)
            )
            QMetaObject.invokeMethod(
                self.startButton,
                'setEnabled',
                Qt.ConnectionType.QueuedConnection,
                Q_ARG(bool, True)
            )
            QMetaObject.invokeMethod(
                self.clearParseSettingsButton,
                'setEnabled',
                Qt.ConnectionType.QueuedConnection,
                Q_ARG(bool, True)
            )


def main() -> None:
    app = QtWidgets.QApplication(sys.argv)

    window = App()
    window.show()

    app.exec()


if __name__ == '__main__':
    main()
