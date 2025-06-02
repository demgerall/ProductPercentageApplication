import json
import logging
import os

from typing import Literal, Any

from PyQt6 import QtWidgets
from PyQt6.QtWidgets import QMessageBox

from tools.constants import AppConstants
from tools.dataConvert import dictToTable, arrayToTable, tableToDict, tableToArray


def _create_default_config(config_type: Literal['app', 'parser'], username: str) -> dict[str, Any]:
    """Создает и сохраняет конфиг по умолчанию для указанного типа.

    Args:
        config_type: Тип конфига ('app' или 'parser')

    Returns:
        Словарь с настройками по умолчанию
    """
    default_configs = {
        'app': {
            'savePath': '',
            'fastExport': 'True',
            'timeDelay': 5
        },
        'parser': {
            'regionCode': 1,
            'requestType': 5,
            'login': '',
            'password': '',
            'isDeliveryDateLimit': 'False',
            'deliveryDateLimit': 1,
            'onlyInStock': 'False',
            'onlyWithGuarantee': 'False',
            'isStoreRatingLimit': 'False',
            'storeRatingLimit': 1,
            'useBlackList': 'False',
            'useWhiteList': 'False',
            'brandsList': {},
            'blackList': [],
            'whiteList': []
        }
    }

    config = default_configs.get(config_type, {})

    config_dir = f'configs/{username}/'
    config_path = os.path.join(config_dir, AppConstants.CONFIG_FILES[config_type])

    os.makedirs(config_dir, exist_ok=True)

    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4)
        return config
    except Exception as ex:
        logging.error(f'Ошибка создания конфига {config_type}: {ex}')
        return {}


def loadAppConfig(window: QtWidgets) -> dict[str, Any]:
    """Загружает и применяет конфигурацию приложения.

    Загружает настройки из конфига приложения и применяет их к UI:
    - Устанавливает путь сохранения
    - Настраивает чекбокс быстрого экспорта
    - Устанавливает задержку между запросами

    Args
        window (QtWidgets.QWidget): Родительское окно для диалоговых сообщений.
            Должно быть виджетом из QtWidgets для корректного отображения QMessageBox.

    Returns:
        dict[str, Any]: Словарь с настройками приложения:
            - savePath (str): Путь для сохранения файлов
            - fastExport (bool): Флаг быстрого экспорта
            - timeDelay (int): Задержка между запросами (сек)

    Side effects:
        - Обновляет placeholder поля standardSavePathInput
        - Устанавливает состояние fastExportCheckBox
        - Устанавливает значение timeDelaySpinBox
        - Обновляет текст statusLabel
    """
    app_config = loadConfig(window, 'app')

    try:
        if app_config.get('savePath'):
            window.standardSavePathInput.setPlaceholderText(app_config.get('savePath'))

        window.fastExportCheckBox.setChecked(
            str(app_config.get('fastExport', 'False')).lower() == 'true'
        )
        window.timeDelaySpinBox.setValue(
            int(app_config.get('timeDelay', 1))
        )

        window.statusLabel.setText('Конфиг приложения успешно загружен')
        return app_config
    except Exception as ex:
        logging.error(f'Ошибка применения конфига приложения: {ex}')
        QMessageBox.warning(window, 'Ошибка', 'Не удалось применить настройки')
        return {}


def loadParserConfig(window: QtWidgets) -> dict[str, Any]:
    """Загружает и применяет конфигурацию парсера.

    Загружает настройки парсера и применяет их к UI:
    - Настраивает параметры доставки и наличия
    - Устанавливает ограничения по рейтингу
    - Применяет черный/белый списки
    - Заполняет соответствующие таблицы

    Args
        window (QtWidgets.QWidget): Родительское окно для диалоговых сообщений.
            Должно быть виджетом из QtWidgets для корректного отображения QMessageBox.

    Returns:
        dict[str, Any]: Словарь с настройками парсера:
            - isDeliveryDateLimit (bool)
            - deliveryDateLimit (int)
            - onlyInStock (bool)
            - и другие параметры...

    Side effects:
        - Обновляет состояние чекбоксов (deliveryDateCheckBox и др.)
        - Устанавливает значения спинбоксов (rateSpinBox и др.)
        - Заполняет таблицы (brandsTable, blackListTable, whiteListTable)
        - Обновляет текст statusLabel
    """
    parser_config = loadConfig(window, 'parser')

    try:
        window.deliveryDateCheckBox.setChecked(
            str(parser_config.get('isDeliveryDateLimit', 'False')).lower() == 'true'
        )
        window.deliveryDateSpinBox.setValue(
            int(parser_config.get('deliveryDateLimit', 1))
        )
        window.instockCheckBox.setChecked(
            str(parser_config.get('onlyInStock', 'False')).lower() == 'true'
        )
        window.guaranteeCheckBox.setChecked(
            str(parser_config.get('onlyWithGuarantee', 'False')).lower() == 'true'
        )
        window.rateCheckBox.setChecked(
            str(parser_config.get('isStoreRatingLimit', 'False')).lower() == 'true'
        )
        window.rateSpinBox.setValue(
            int(parser_config.get('storeRatingLimit', 1))
        )
        window.blackListCheckBox.setChecked(
            str(parser_config.get('useBlackList', 'False')).lower() == 'true'
        )
        window.whiteListCheckBox.setChecked(
            str(parser_config.get('useWhiteList', 'False')).lower() == 'true'
        )
        dictToTable(parser_config.get('brandsList', {}), window.brandsTable)
        arrayToTable(parser_config.get('blackList', []), window.blackListTable)
        arrayToTable(parser_config.get('whiteList', []), window.whiteListTable)

        window.statusLabel.setText('--Загрузка конфига парсера прошла успешно--')
        return parser_config
    except Exception as ex:
        logging.error(f'Ошибка применения конфига парсера: {ex}')
        QMessageBox.warning(window, 'Ошибка', 'Не удалось применить настройки парсера')
        return {}


def loadConfig(window: QtWidgets, config_type: Literal['app', 'parser']) -> dict[str, Any]:
    """Загружает конфигурацию из JSON-файла и возвращает как словарь.

    Args:
        window (QtWidgets.QWidget): Родительское окно для диалоговых сообщений.
            Должно быть виджетом из QtWidgets для корректного отображения QMessageBox.
        config_type (Literal['app', 'parser']): Тип загружаемого конфига:
            - 'app' - конфиг приложения
            - 'parser' - конфиг парсера

    Returns:
        dict[str, Any]: Словарь с настройками конфигурации. Структура зависит от типа:
            - Для 'app': содержит savePath, fastExport, timeDelay
            - Для 'parser': содержит isDeliveryDateLimit, deliveryDateLimit и др.

    Raises:
        ValueError: Если передан неверный config_type
        FileNotFoundError: Если файл конфига не существует
        JSONDecodeError: Если файл содержит невалидный JSON

    Note:
        В случае ошибки показывает сообщение QMessageBox и возвращает пустой словарь
    """
    if config_type not in AppConstants.CONFIG_FILES:
        raise ValueError(f'Неподдерживаемый тип конфига: {config_type}')

    config_path = f'configs/{window.username}/' + AppConstants.CONFIG_FILES[config_type]
    if not os.path.exists(config_path):
        return _create_default_config(config_type, window.username)

    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except json.JSONDecodeError:
        logging.warning(f'Невалидный JSON в {config_path}, создаю новый конфиг')
        return _create_default_config(config_type, window.username)
    except Exception as ex:
        logging.error(f'Ошибка загрузки {config_path}: {ex}')
        QMessageBox.critical(window, 'Ошибка', f'Не удалось загрузить {config_type} конфиг')
        return _create_default_config(config_type, window.username)


def saveAppConfig(window: QtWidgets) -> None:
    """Сохраняет текущие настройки приложения в конфигурационный файл.

    Собирает данные из UI элементов и сохраняет их в конфиг приложения:
    - Путь для сохранения файлов (standardSavePathInput)
    - Настройку быстрого экспорта (fastExportCheckBox)
    - Задержку между запросами (timeDelaySpinBox)

    Args
        window (QtWidgets.QWidget): Родительское окно для диалоговых сообщений.
            Должно быть виджетом из QtWidgets для корректного отображения QMessageBox.

    Returns:
        None: Сохраняет изменения в файл через saveConfig()

    Side effects:
        - Обновляет placeholder для standardSavePathInput
        - Очищает текст в standardSavePathInput после сохранения
        - Устанавливает статус в statusLabel

    Note:
        Сохраняет только измененные параметры для минимизации операций записи
    """
    current_config = {
        'savePath': window.standardSavePathInput.text().strip() or window.app_config['savePath'],
        'fastExport': str(window.fastExportCheckBox.isChecked()),
        'timeDelay': window.timeDelaySpinBox.value()
    }

    if current_config != window.app_config:
        window.app_config.update(current_config)

        if current_config['savePath'] != window.standardSavePathInput.placeholderText():
            window.standardSavePathInput.setPlaceholderText(current_config['savePath'])
            window.standardSavePathInput.clear()

        saveConfig(window, window.app_config, 'app')
        window.statusLabel.setText('Настройки приложения сохранены')


def saveParserConfig(window: QtWidgets) -> None:
    """Сохраняет текущие настройки парсера в конфигурационный файл.

    Собирает данные из всех связанных UI элементов:
    - Настройки доставки и наличия
    - Ограничения по рейтингу магазинов
    - Использование черного/белого списков
    - Данные из таблиц (brandsList, blackList, whiteList)

    Args
        window (QtWidgets.QWidget): Родительское окно для диалоговых сообщений.
            Должно быть виджетом из QtWidgets для корректного отображения QMessageBox.

    Returns:
        None: Сохраняет изменения в файл через saveConfig()

    Side effects:
        - Обновляет данные в таблицах (brandsTable, blackListTable, whiteListTable)
        - Устанавливает статус в statusLabel
    """
    new_config = {
        'regionCode': 1,
        'requestType': 5,
        'login': '',
        'password': '',
        'isDeliveryDateLimit': str(window.deliveryDateCheckBox.isChecked()),
        'deliveryDateLimit': window.deliveryDateSpinBox.value(),
        'onlyInStock': str(window.instockCheckBox.isChecked()),
        'onlyWithGuarantee': str(window.guaranteeCheckBox.isChecked()),
        'isStoreRatingLimit': str(window.rateCheckBox.isChecked()),
        'storeRatingLimit': window.rateSpinBox.value(),
        'useBlackList': str(window.blackListCheckBox.isChecked()),
        'useWhiteList': str(window.whiteListCheckBox.isChecked()),
        'brandsList': tableToDict(window.brandsTable),
        'blackList': tableToArray(window.blackListTable),
        'whiteList': tableToArray(window.whiteListTable)
    }

    if new_config != window.parser_config:
        window.parser_config.update(new_config)
        saveConfig(window, window.parser_config, 'parser')
        window.statusLabel.setText('Настройки парсера сохранены')


def saveConfig(window: QtWidgets, config: dict[str, Any], config_type: Literal['app', 'parser']) -> None:
    """Сохраняет конфигурацию в соответствующий JSON-файл.

    Args:
        window (QtWidgets.QWidget): Родительское окно для диалоговых сообщений.
            Должно быть виджетом из QtWidgets для корректного отображения QMessageBox.
        config (dict): Словарь с настройками для сохранения
        config_type (Literal['app', 'parser']): Тип конфигурации

    Raises:
        PermissionError: При отсутствии прав на запись файла
        IOError: При проблемах с записью на диск

    Note:
        В случае ошибки показывает QMessageBox с описанием проблемы
        и сохраняет детали в лог
    """
    config_path = f'configs/{window.username}/' + AppConstants.CONFIG_FILES[config_type]

    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
        logging.info(f'Конфиг {config_type} успешно сохранён')
    except PermissionError:
        error_msg = f'Нет прав для записи в {config_path}'
        logging.error(error_msg)
        QMessageBox.critical(window, 'Ошибка сохранения', error_msg)
    except Exception as ex:
        error_msg = f'Ошибка сохранения {config_type}: {str(ex)}'
        logging.exception(error_msg)
        QMessageBox.critical(
            window,
            'Ошибка сохранения',
            f'Не удалось сохранить {config_type} конфиг\n{str(ex)}'
        )
