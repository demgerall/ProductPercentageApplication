import logging
import requests

from typing import Optional

from PyQt6 import QtWidgets

from tools.constants import AppConstants
from tools.XMLToDict import parseXMLResponseToDict


def safeAPIRequest(window: QtWidgets, params: dict) -> Optional[dict]:
    """
    Выполняет безопасный запрос к API с обработкой возможных ошибок.

    Args:
        window (QtWidgets.QWidget): Родительское окно для диалоговых сообщений.
            Должно быть виджетом из QtWidgets для корректного отображения QMessageBox.
        params (dict): Параметры запроса, которые будут переданы в GET-запросе.

    Returns:
        Optional[dict]: Словарь с данными ответа в случае успеха, None в случае ошибки.

    Note:
        - Используется стандартный таймаут из AppConstants.API_TIMEOUT
        - Включена верификация SSL сертификата (verify=True)
        - Обрабатываются следующие исключения:
          * requests.RequestException - проблемы с сетевым запросом
          * ValueError - проблемы при парсинге XML ответа
        - Все ошибки логируются с указанием деталей исключения
    """
    try:
        response = requests.get(
            url=window.api_url,
            params=params,
            timeout=AppConstants.API_TIMEOUT,
            verify=True,
            headers={'User-Agent': 'Mozilla/5.0'}
        )

        response.raise_for_status()

        if not response.content:
            logging.warning('Получен пустой ответ от API')
            return

        return parseXMLResponseToDict(response)

    except requests.Timeout:
        logging.error(f'Таймаут соединения с API (превышено {AppConstants.API_TIMEOUT} секунд)')
        return

    except requests.HTTPError as http_err:
        status_code = response.status_code if 'response' in locals() else 'неизвестен'
        logging.error(f'Ошибка HTTP {status_code}: {str(http_err)}')
        return

    except requests.ConnectionError:
        logging.error('Ошибка подключения к API: невозможно установить соединение')
        return

    except (requests.RequestException, ValueError) as ex:
        logging.error(f'Ошибка при выполнении запроса к API: {str(ex)}', exc_info=True)
        return
