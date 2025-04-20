import logging

from PyQt6 import QtWidgets
from PyQt6.QtWidgets import QTableWidgetItem, QTableWidget, QMessageBox


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


def removeTableRow(window: QtWidgets.QWidget, table: QtWidgets.QTableWidget) -> None:
    """Удаляет выбранные строки из таблицы после подтверждения действия.

    Args:
        window (QtWidgets.QWidget): Родительское окно для диалоговых сообщений.
            Должно быть виджетом из QtWidgets для корректного отображения QMessageBox.
        table (QtWidgets.QTableWidget): Таблица QTableWidget, из которой будут
            удаляться строки. Должна содержать хотя бы одну строку.

    Raises:
        TypeError: Если переданные аргументы не являются Qt виджетами.
        ValueError: Если таблица пуста или не содержит выделенных строк.

    Note:
        - Удаление происходит в обратном порядке для сохранения корректности индексов.
        - Перед удалением запрашивается подтверждение у пользователя.
        - Показывает предупреждение, если не выбрано ни одной строки.

    Examples:
        >>> # Удаление строк из таблицы с подтверждением
        >>> removeTableRow(main_window, data_table)
    """
    if not isinstance(table, QtWidgets.QTableWidget):
        raise TypeError('Аргумент table должен быть QTableWidget')

    if table.rowCount() == 0:
        raise ValueError('Таблица не содержит строк для удаления')

    selected_rows = {item.row() for item in table.selectedItems()}

    if not selected_rows:
        QMessageBox.warning(
            window,
            'Не выбраны строки',
            'Пожалуйста, выделите строки для удаления'
        )
        return

    confirm = QMessageBox.question(
        window,
        'Подтверждение удаления',
        f'Вы действительно хотите удалить {len(selected_rows)} строк(у)?',
        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        QMessageBox.StandardButton.No
    )

    if confirm == QMessageBox.StandardButton.No:
        return

    for row in sorted(selected_rows, reverse=True):
        if 0 <= row < table.rowCount():
            table.removeRow(row)
        else:
            logging.warning(f'Попытка удаления несуществующей строки: {row}')


def validateTable(window: QtWidgets, table: QtWidgets.QTableWidget) -> bool:
    """Проверяет таблицу на наличие пустых ячеек и удаляет строки с ними после подтверждения.

    Args:
        window (QtWidgets.QWidget): Родительское окно для диалоговых сообщений.
            Должен быть виджетом из QtWidgets для корректного отображения QMessageBox.
        table (QtWidgets.QTableWidget): Таблица QTableWidget для валидации.

    Returns:
        bool:
            - True если таблица валидна (нет пустых ячеек) или пользователь подтвердил удаление
            - False если пользователь отменил операцию удаления строк

    Raises:
        TypeError: Если аргументы не являются Qt виджетами.
        ValueError: Если таблица не содержит строк.

    Note:
        - Пустыми считаются ячейки: None, с пустым текстом или только с пробелами
        - Удаление происходит в обратном порядке для сохранения корректности индексов
        - Перед удалением запрашивается подтверждение у пользователя

    Examples:
        >>> if validateTable(main_window, data_table):
        ...     print("Валидация прошла успешно")
    """
    if not isinstance(table, QtWidgets.QTableWidget):
        raise TypeError("Аргумент table должен быть QTableWidget")

    row_count = table.rowCount()
    if row_count == 0:
        return True

    rows_to_remove = set()
    for row in range(row_count):
        for col in range(table.columnCount()):
            item = table.item(row, col)
            if item is None or not item.text().strip():
                rows_to_remove.add(row)
                break

    if not rows_to_remove:
        return True

    reply = QMessageBox.question(
        window,
        'Пустые значения в таблице',
        f'Найдено {len(rows_to_remove)} строк с пустыми значениями. Удалить эти строки?',
        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        QMessageBox.StandardButton.No
    )

    if reply == QMessageBox.StandardButton.No:
        return False

    for row in sorted(rows_to_remove, reverse=True):
        if 0 <= row < row_count:
            table.removeRow(row)

    return True
