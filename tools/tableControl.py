from PyQt6.QtWidgets import QTableWidgetItem, QTableWidget


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
        raise ValueError('Таблица должна содержать минимум 2 колонки')
    if columns > table.columnCount():
        raise ValueError(
            f'Параметр columns ({columns}) превышает количество колонок в таблице '
            f'({table.columnCount()})'
        )

    row = table.rowCount()
    table.insertRow(row)

    for col in range(columns):
        table.setItem(row, col, QTableWidgetItem(''))