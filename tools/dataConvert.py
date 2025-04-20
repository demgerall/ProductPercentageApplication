import pandas as pd

from typing import Any

from PyQt6.QtWidgets import QTableWidget, QTableWidgetItem


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
            - Пустые/несуществующие ячейки представлены пустой строкой ''

    Raises:
        TypeError: Если передан не QTableWidget объект

    Examples:
        >>> table = QTableWidget(2, 3)  # Создаем таблицу 2x3
        >>> table.setItem(0, 0, QTableWidgetItem('Текст'))
        >>> tableToArray(table)
        [['Текст', '', ''], ['', '', '']]

        >>> # Пустая таблица
        >>> table.setRowCount(0)
        >>> tableToArray(table)
        []
    """
    if not isinstance(table, QTableWidget):
        raise TypeError(f'Ожидается QTableWidget, получен {type(table).__name__}')

    result = []

    for row in range(table.rowCount()):
        row_data = []

        for col in range(table.columnCount()):
            item = table.item(row, col)
            text = item.text().strip() if item is not None else ''
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
        >>> data = [['A1', 'B1'], ['A2', 'B2']]
        >>> arrayToTable(data, table)
        >>> table.rowCount()
        2
        >>> table.columnCount()
        2

        >>> # Обработка None значений
        >>> arrayToTable([['X', None], [None, 'Y']], table)
        >>> table.item(0, 1).text()
        ''
    """
    if not isinstance(table, QTableWidget):
        raise TypeError(f'Ожидается QTableWidget, получен {type(table).__name__}')
    if not all(isinstance(row, list) for row in data):
        raise TypeError('Data должен быть двумерным списком')
    if data and len({len(row) for row in data}) > 1:
        raise ValueError('Все строки в data должны иметь одинаковую длину')

    table.setRowCount(0)

    if not data:
        return

    table.setRowCount(len(data))
    table.setColumnCount(len(data[0]))

    for row_idx, row_data in enumerate(data):
        for col_idx, cell_data in enumerate(row_data):
            item = QTableWidgetItem(str(cell_data) if cell_data is not None else '')
            table.setItem(row_idx, col_idx, item)


def tableToDict(table: QTableWidget) -> dict[str, str]:
    """Преобразует первые две колонки QTableWidget в словарь пар ключ-значение.

    Парсит таблицу построчно, используя значения из первого столбца как ключи,
    а из второго столбца - как соответствующие значения. Пустые ячейки и None
    значения преобразуются в пустые строки. Строки с пустыми ключами пропускаются.

    Args:
        table (QTableWidget): Исходная таблица с минимум двумя колонками.
            Должна содержать данные в первых двух колонках.

    Returns:
        dict[str, str]: Словарь, где:
            - Ключи: значения из первого столбца (колонка 0)
            - Значения: соответствующие значения из второго столбца (колонка 1)
            - Все ключи и значения преобразуются в строки

    Raises:
        ValueError: Если в таблице меньше двух колонок
        TypeError: Если передан не QTableWidget объект

    Examples:
        >>> table = QTableWidget(2, 2)
        >>> table.setItem(0, 0, QTableWidgetItem('key1'))
        >>> table.setItem(0, 1, QTableWidgetItem('value1'))
        >>> table.setItem(1, 0, QTableWidgetItem('key2'))
        >>> tableToDict(table)
        {'key1': 'value1', 'key2': ''}

        >>> # Пропуск строк с пустыми ключами
        >>> table.setItem(1, 0, QTableWidgetItem(''))
        >>> tableToDict(table)
        {'key1': 'value1'}
    """
    if not isinstance(table, QTableWidget):
        raise TypeError(f'Ожидается QTableWidget, получен {type(table).__name__}')
    if table.columnCount() < 2:
        raise ValueError('Таблица должна содержать минимум 2 колонки')

    result = {}

    for row in range(table.rowCount()):
        key_item = table.item(row, 0)
        key = key_item.text().strip() if key_item is not None else ""

        value_item = table.item(row, 1)
        value = value_item.text().strip() if value_item is not None else ""

        if key:
            result[key] = value

    return result


def dictToTable(data: dict[str, Any], table: QTableWidget) -> None:
    """Заполняет QTableWidget данными из словаря, преобразуя пары ключ-значение в строки таблицы.

    Преобразует словарь в табличное представление, где:
    - Ключи словаря размещаются в первом столбце (колонка 0)
    - Значения словаря размещаются во втором столбце (колонка 1)
    - None значения преобразуются в пустые строки
    - Все нестроковые значения автоматически конвертируются в строки

    Args:
        data (dict[str, Any]): Входной словарь для преобразования, где:
            - Ключи: будут помещены в первый столбец таблицы
            - Значения: будут помещены во второй столбец таблицы
        table (QTableWidget): Целевая таблица для заполнения. Существующие данные
            будут полностью заменены. Таблица должна иметь как минимум 2 колонки.

    Returns:
        None: Метод модифицирует переданную таблицу напрямую.

    Raises:
        TypeError: Если передан не QTableWidget объект
        ValueError: Если таблица содержит меньше 2 колонок

    Examples:
        >>> table = QTableWidget()
        >>> data = {'Key1': 'Value1', 'Key2': 123, 'Key3': None}
        >>> dictToTable(data, table)
        >>> table.rowCount()
        3
        >>> table.item(0, 0).text()
        'Key1'
        >>> table.item(1, 1).text()
        '123'
        >>> table.item(2, 1).text()
        ''
    """
    if not isinstance(table, QTableWidget):
        raise TypeError(f'Ожидается QTableWidget, получен {type(table).__name__}')
    if table.columnCount() < 2:
        raise ValueError('Таблица должна содержать минимум 2 колонки')

    table.setRowCount(len(data))

    for row, (key, value) in enumerate(data.items()):
        key_item = QTableWidgetItem(str(key))
        table.setItem(row, 0, key_item)

        value_item = QTableWidgetItem(str(value) if value is not None else "")
        table.setItem(row, 1, value_item)


def tableFromDataframe(table: QTableWidget, data: pd.DataFrame) -> None:
    """Заполняет QTableWidget данными из pandas DataFrame с автоматической настройкой заголовков.

    Полностью заменяет содержимое таблицы данными из DataFrame, включая:
    - Установку соответствующего количества строк и столбцов
    - Перенос заголовков столбцов DataFrame
    - Автоматическое преобразование значений в строки
    - Обработку NaN/None значений (преобразуются в пустые строки)
    - Автоматическую подгонку ширины столбцов под содержимое

    Args:
        table (QTableWidget): Целевая таблица Qt для заполнения. Существующее содержимое
            будет полностью очищено перед заполнением.
        data (pd.DataFrame): DataFrame для переноса в таблицу. Должен содержать:
            - Заголовки столбцов (для переноса в горизонтальные заголовки таблицы)
            - Данные, поддерживающие преобразование в строки

    Returns:
        None: Метод модифицирует переданную таблицу напрямую.

    Raises:
        TypeError: Если входные данные не являются pandas DataFrame
        ValueError: Если DataFrame пуст или содержит несовместимые данные

    Examples:
        >>> df = pd.DataFrame({
        ...     'Name': ['Alice', 'Bob', None],
        ...     'Age': [25, 30, pd.NA],
        ...     'Score': [4.5, 3.8, 5.0]
        ... })
        >>> table = QTableWidget()
        >>> tableFromDataframe(table, df)
        >>> table.rowCount()
        3
        >>> table.columnCount()
        3
        >>> table.horizontalHeaderItem(0).text()
        'Name'
        >>> table.item(2, 0).text()
        ''
    """
    if not isinstance(data, pd.DataFrame):
        raise TypeError(f'Ожидается pandas DataFrame, получен {type(data).__name__}')
    if data.empty:
        raise ValueError('DataFrame не должен быть пустым')

    table.clear()

    n_rows, n_cols = data.shape
    table.setRowCount(n_rows)
    table.setColumnCount(n_cols)

    table.setHorizontalHeaderLabels(data.columns.tolist())

    for i in range(n_rows):
        for j in range(n_cols):
            value = data.iloc[i, j]
            item_value = str(value) if not pd.isna(value) else ""
            item = QTableWidgetItem(item_value)
            table.setItem(i, j, item)

    table.resizeColumnsToContents()