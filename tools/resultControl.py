from typing import Any

from PyQt6 import QtWidgets

from tools.constants import AppConstants


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


def validateResult(window: QtWidgets, response_data_table: list[dict]) -> list[dict]:
    """
    Фильтрует результаты парсинга согласно заданным в конфигурации правилам.

    Args:
        window (QtWidgets.QWidget): Родительское окно для диалоговых сообщений.
            Должно быть виджетом из QtWidgets для корректного отображения QMessageBox.
        response_data_table (list[dict]): Список словарей с данными для фильтрации, где каждый словарь содержит:
            - delivery_days: int - срок доставки
            - instock: int - наличие товара (1 - в наличии)
            - rating: float - рейтинг магазина
            - class_man: str - идентификатор производителя
            - class_user: str - идентификатор пользователя

    Returns:
        list[dict]: Отфильтрованный список словарей, соответствующий всем условиям фильтрации

    Note:
        Применяются следующие фильтры (если включены в конфигурации):
        1. Ограничение по сроку доставки
        2. Только товары в наличии
        3. Ограничение по рейтингу магазина
        4. Черный список производителей/пользователей
        5. Белый список производителей/пользователей
    """
    results = []
    config = window.parser_config

    for item in response_data_table:
        if config.get('isDeliveryDateLimit') == 'True':
            if item['delivery_days'] > config['deliveryDateLimit']:
                continue

        if config.get('onlyInStock') == 'True' and item['instock'] != 1:
            continue

        if config.get('onlyWithGuarantee') == 'True' and item['descr_qtyV2'].lower().find('гарантия') == -1:
            continue

        if config.get('isStoreRatingLimit') == 'True':
            if item['rating'] < config['storeRatingLimit']:
                continue

        if config.get('useBlackList') == 'True' and config['blackList']:
            item_pair = [item['class_man'], item['class_user']]
            if item_pair in config['blackList']:
                continue

        if config.get('useWhiteList') == 'True' and config['whiteList']:
            item_pair = [item['class_man'], item['class_user']]
            if item_pair not in config['whiteList']:
                continue

        results.append(item)

    return results


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
            str(int(row['priceV2'])),
            str(row['qtyV2'] if row['qtyV2'] >= 0 else 0),
            str(row['descr_qtyV2']),
            str(row['class_cat']),
            str(row['class_user']),
            str(row['descr_price']),
            str(row['delivery_days'])
        ])

    return result_data_row
