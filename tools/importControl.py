import logging
import os

import pandas as pd

from PyQt6 import QtWidgets
from PyQt6.QtWidgets import QMessageBox, QTableWidgetItem, QFileDialog, QTableWidget

from tools.constants import AppConstants


def loadSearchExcelFilePath(window: QtWidgets) -> None:
    """Загружает Excel-файл с артикулами для последующего парсинга.

    Открывает диалоговое окно выбора файла и обрабатывает выбранный файл:
    1. Позволяет пользователю выбрать файл Excel (.xlsx)
    2. Проверяет, что файл был выбран
    3. Сохраняет путь к файлу в search_file_path_Excel
    4. Обновляет интерфейс, отображая имя выбранного файла

    Args:
        window (QtWidgets.QWidget): Родительское окно для диалоговых сообщений.
            Должно быть виджетом из QtWidgets для корректного отображения QMessageBox.

    Side effects:
        - Устанавливает значение search_file_path_Excel
        - Обновляет текст choosedFileLabel
        - Может изменить состояние других элементов UI

    Raises:
        ValueError: Если выбранный файл имеет недопустимое расширение
    """
    file_path, _ = QFileDialog.getOpenFileName(
        parent=window,
        caption='Выберите файл с артикулами',
        directory='',
        filter='Excel Files (*.xlsx);;All Files (*)'
    )

    if not file_path:
        window.choosedFileLabel.setText('Файл не выбран')
        return

    if not file_path.lower().endswith('.xlsx'):
        QMessageBox.warning(
            window,
            'Неверный формат файла',
            'Пожалуйста, выберите файл в формате Excel (.xlsx)'
        )
        return

    window.search_file_path_Excel = file_path
    file_name = os.path.basename(file_path)
    window.choosedFileLabel.setText(file_name)
    window.statusLabel.setText(f'Выбран файл: {file_name}')


def importSearchExcelFileToArray(window: QtWidgets, path: str) -> list[list[str]] | None:
    """Загружает и валидирует данные из Excel-файла с артикулами для поиска.

    Функция выполняет следующие действия:
    1. Загружает данные из Excel-файла по указанному пути
    2. Проверяет наличие данных в файле
    3. Валидирует структуру таблицы (наличие требуемых колонок)
    4. Очищает данные от пустых значений с опциональным подтверждением пользователя
    5. Возвращает данные в виде списка списков строк или None в случае ошибки

    Args:
        window (QtWidgets.QWidget): Родительское окно для диалоговых сообщений.
            Должно быть виджетом из QtWidgets для корректного отображения QMessageBox.
        path (str): Путь к Excel-файлу для импорта

    Returns:
        list[list[str]] | None: Список строк с данными (каждая строка - список значений),
                                или None если возникла ошибка или пользователь отменил импорт

    Raises:
        Exception: Логирует любые исключения при работе с файлом, но не пробрасывает их выше
    """
    try:
        df = pd.read_excel(path)

        if df.empty:
            QMessageBox.warning(
                window,
                'Нет данных для импорта',
                'Импортируемый файл не содержит данных'
            )
            return None

        required_columns = AppConstants.COLUMNS['SEARCH']
        if list(df.columns) != required_columns:
            QMessageBox.warning(
                window,
                'Ошибка формата импортируемой таблицы',
                f'Импортируемый файл должен содержать колонки: {", ".join(required_columns)}'
            )
            return None

        if df.isnull().any().any() or any(df.apply(lambda row: row.str.strip().eq('').any(), axis=1)):
            df = df.dropna(how='any').reset_index(drop=True)
            df = df[~df.apply(lambda row: row.str.strip().eq('').any(), axis=1)].reset_index(drop=True)

            if df.empty:
                QMessageBox.warning(
                    window,
                    'Нет данных для импорта',
                    'После удаления пустых строк в таблице не осталось данных'
                )
                return None

            reply = QMessageBox.question(
                window,
                'Обнаружены пустые значения',
                'В таблице обнаружены пустые ячейки или строки. Продолжить импорт? (Пустые строки будут удалены)',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )

            if reply == QMessageBox.StandardButton.No:
                return None

        return df.values.tolist()

    except Exception as ex:
        logging.exception("Ошибка при импорте Excel-файла", exc_info=ex)
        QMessageBox.critical(
            window,
            'Ошибка импорта',
            'Не удалось загрузить файл. Проверьте формат файла и попробуйте снова.'
        )
        return None


def importListExcelFile(window: QtWidgets, table: QTableWidget) -> None:
    """
    Загружает данные из Excel файла в QTableWidget для Черного/Белого списка с валидацией.

    Args:
        window (QtWidgets.QWidget): Родительское окно для диалоговых сообщений.
            Должно быть виджетом из QtWidgets для корректного отображения QMessageBox.
        table (QTableWidget): Целевая таблица для загрузки данных

    Raises:
        - Отсутствие файла или отмена выбора
        - Несоответствие формата (заголовки столбцов)
        - Пустые данные после валидации
        - Ошибки чтения файла
        - Проблемы с доступом к файлу

    Note:
        1. Открывает диалог выбора файла Excel
        2. Проверяет наличие данных в файле
        3. Валидирует структуру файла (заголовки столбцов)
        4. Удаляет пустые строки и ячейки
        5. Запрашивает подтверждение при наличии пустых значений
        6. Загружает данные в таблицу Qt
        7. Выводит результат операции
    """
    file_path, _ = QFileDialog.getOpenFileName(
        window,
        'Выберите файл Excel со списком',
        '',
        'Excel Files (*.xlsx);;Все файлы (*)'
    )

    if not file_path:
        return

    try:
        df = pd.read_excel(file_path, dtype=str)

        if df.empty:
            QMessageBox.warning(
                window,
                'Пустой файл',
                'Выбранный файл не содержит данных'
            )
            return

        required_columns = AppConstants.COLUMNS['LISTS']
        if list(df.columns) != required_columns:
            QMessageBox.warning(
                window,
                'Неверный формат',
                f'Файл должен содержать 2 столбца с заголовками: {", ".join(required_columns)}'
            )
            return

        if df.isnull().any().any() or any(df.apply(lambda row: row.str.strip().eq('').any(), axis=1)):
            df = df.dropna(how='any').reset_index(drop=True)
            df = df[~df.apply(lambda row: row.str.strip().eq('').any(), axis=1)].reset_index(drop=True)

            if df.empty:
                QMessageBox.warning(
                    window,
                    'Нет данных для импорта',
                    'После удаления пустых строк в таблице не осталось данных'
                )
                return

            reply = QMessageBox.question(
                window,
                'Обнаружены пустые значения',
                f'Строки с пустыми значениями будут удалены. Продолжить импорт?',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                return

        table.setRowCount(0)
        table.setRowCount(len(df))

        for row in range(len(df)):
            for col in range(2):
                value = str(df.iloc[row, col]).strip()
                item = QTableWidgetItem(value)
                table.setItem(row, col, item)

        table.resizeColumnsToContents()

        QMessageBox.information(
            window,
            'Импорт завершен',
            f'Успешно импортировано {len(df)} строк\nФайл: {os.path.basename(file_path)}'
        )

    except PermissionError:
        QMessageBox.critical(
            window,
            'Ошибка доступа',
            'Невозможно прочитать файл. Закройте файл если он открыт.'
        )
    except Exception as ex:
        logging.error(f'Ошибка импорта: {str(ex)}', exc_info=True)
        QMessageBox.critical(
            window,
            'Ошибка импорта',
            f'Не удалось загрузить файл:\n{str(ex)}'
        )
