import logging
import datetime

import pandas as pd

from typing import Literal

from PyQt6 import QtWidgets
from PyQt6.QtWidgets import QMessageBox, QTableWidget, QFileDialog

from tools.constants import AppConstants


def exportListExcelFile(window: QtWidgets, table: QTableWidget, table_type: Literal['black', 'white']) -> None:
    """
    Экспортирует данные из QTableWidget в Excel файл с предварительной валидацией.

    Args:
        window (QtWidgets.QWidget): Родительское окно для диалоговых сообщений.
            Должно быть виджетом из QtWidgets для корректного отображения QMessageBox.
        table (QTableWidget): Таблица Qt, содержащая данные для экспорта
        table_type: Тип таблицы ('black' или 'white')

    Raises:
        - Отсутствие данных в таблице
        - Пустые строки после валидации
        - Ошибки сохранения файла
        - Проблемы с доступом к файловой системе

    Note:
        1. Проверяет наличие данных в таблице
        2. Собирает данные, пропуская пустые ячейки
        3. Проверяет валидность собранных данных (должны быть ровно 2 столбца)
        4. Запрашивает подтверждение если найдены пустые ячейки
        5. Открывает диалог сохранения файла
        6. Сохраняет данные в Excel
        7. Выводит результат операции
    """
    if table.rowCount() == 0:
        QMessageBox.warning(
            window,
            'Нет данных для экспорта',
            'Экспортируемая таблица не содержит данных'
        )
        return

    headers = AppConstants.COLUMNS['LISTS']
    valid_data = []
    empty_rows_count = 0

    for row in range(table.rowCount()):
        row_data = []

        for col in range(table.columnCount()):
            item = table.item(row, col)

            if item is None or not item.text().strip():
                continue

            row_data.append(item.text().strip())

        if len(row_data) == 2:
            valid_data.append(row_data)
        else:
            empty_rows_count += 1

    if not valid_data:
        QMessageBox.warning(
            window,
            'Нет данных для экспорта',
            'После удаления строк с пустыми ячейками таблица стала пустой. Экспорт отменен'
        )
        return

    if empty_rows_count > 0:
        reply = QMessageBox.question(
            window,
            'Обнаружены пустые значения',
            f'Найдено {empty_rows_count} строк с пустыми значениями. Продолжить экспорт? '
            '(Пустые строки будут пропущены)',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.No:
            return

    file_path, _ = QFileDialog.getSaveFileName(
        window,
        'Сохранить список как Excel',
        f'{"Черный" if table_type == "black" else "Белый"}_cписок_{datetime.datetime.now().strftime("%Y-%m-%d")}.xlsx',
        'Excel Files (*.xlsx)'
    )

    if not file_path:
        return

    try:
        df = pd.DataFrame(valid_data, columns=headers)

        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)

            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            header_format = workbook.add_format({
                'bold': True,
                'border': 1,
                'bg_color': '#607ebc',
                'font_color': '#faf5ee',
                'align': 'center'
            })

            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)

            for i, column in enumerate(df.columns):
                try:
                    str_lengths = df[column].astype(str).str.len()
                    max_len = max(str_lengths.max(), len(column))
                    width = min(50, (max_len + 2) * 1.1)
                    worksheet.set_column(i, i, width)
                except AttributeError:
                    logging.warning(f"Ошибка в столбце {column}: {str(e)}")
                    continue
                except Exception as e:
                    logging.warning(f"Ошибка в столбце {column}: {str(e)}")
                    continue

            worksheet.freeze_panes(1, 0)

        QMessageBox.information(
            window,
            'Экспорт завершен',
            f'Данные успешно экспортированы в файл:\n{file_path}'
        )

    except PermissionError:
        QMessageBox.critical(
            window,
            'Ошибка доступа',
            'Невозможно сохранить файл. Закройте файл если он открыт.'
        )
    except Exception as ex:
        logging.error(f'Ошибка экспорта: {str(ex)}', exc_info=True)
        QMessageBox.critical(
            window,
            'Ошибка экспорта',
            f'Не удалось экспортировать данные:\n{str(ex)}'
        )


def exportErrorArticlesExcelFile(window: QtWidgets.QWidget, data: pd.DataFrame) -> None:
    """
    Экспортирует DataFrame с ошибочными артикулами в Excel файл с предварительной проверкой данных.

    Args:
        window (QtWidgets): Родительское окно для диалоговых сообщений.
        data (pd.DataFrame): DataFrame для экспорта. Если пустой, функция отменяется без оповещения.

    Raises:
        - Ошибки сохранения файла
        - Проблемы с доступом к файловой системе
    """
    if data.empty:
        return

    QMessageBox.information(
        window,
        'Экспорт ошибочных артикулов',
        'Проводится экспорт ошибочных артикулов'
    )

    file_path, _ = QFileDialog.getSaveFileName(
        window,
        'Сохранить список ошибочных артикулов',
        f'Ошибочные_артикулы_{datetime.datetime.now().strftime("%Y-%m-%d")}.xlsx',
        'Excel Files (*.xlsx)'
    )

    if not file_path:
        return

    try:
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            data.to_excel(writer, index=False)

            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            header_format = workbook.add_format({
                'bold': True,
                'border': 1,
                'bg_color': '#607ebc',
                'font_color': '#faf5ee',
                'align': 'center'
            })

            for col_num, value in enumerate(data.columns.values):
                worksheet.write(0, col_num, value, header_format)

            for i, column in enumerate(data.columns):
                try:
                    str_lengths = data[column].astype(str).str.len()
                    max_len = max(str_lengths.max(), len(column))
                    width = min(50, (max_len + 2) * 1.1)

                    worksheet.set_column(i, i, width)
                except AttributeError:
                    logging.warning(f"Ошибка в столбце {column}: {str(e)}")
                    continue
                except Exception as e:
                    logging.warning(f"Ошибка в столбце {column}: {str(e)}")
                    continue

            worksheet.freeze_panes(1, 0)

        QMessageBox.information(
            window,
            'Экспорт завершен',
            f'Данные успешно экспортированы в файл:\n{file_path}'
        )

    except PermissionError:
        QMessageBox.critical(
            window,
            'Ошибка доступа',
            'Невозможно сохранить файл. Закройте файл если он открыт.'
        )
    except Exception as ex:
        logging.error(f'Ошибка экспорта: {str(ex)}', exc_info=True)
        QMessageBox.critical(
            window,
            'Ошибка экспорта',
            f'Не удалось экспортировать данные:\n{str(ex)}'
        )


def exportResultExcelFile(window: QtWidgets, save_type: str) -> None:
    """
    Экспортирует результаты парсинга в Excel файл с форматированием.

    Args:
        window (QtWidgets.QWidget): Родительское окно для диалоговых сообщений.
            Должно быть виджетом из QtWidgets для корректного отображения QMessageBox.
        save_type (str): Тип сохранения:
            - 'standard' - сохраняет в стандартную папку (self.base_save_path или из конфига)
            - любое другое значение - открывает диалог выбора файла

    Raises:
        - Проблемы с созданием файла
        - Ошибки записи данных
        - Проблемы с форматами

    Note:
        1. Формирует имя файла с текущей датой/временем
        2. Определяет путь сохранения в зависимости от save_type
        3. Создает Excel файл с помощью xlsxwriter
        4. Применяет форматирование:
            - Заголовки столбцов с разными стилями
            - Разные форматы для числовых данных
            - Особое форматирование для отсутствующих данных
            - Автоподбор ширины столбцов
            - Закрепление заголовков
        5. Обрабатывает ошибки экспорта
    """
    file_name = f'Проценка товара от {datetime.datetime.now().strftime("%d-%b-%Y %H-%M-%S")}.xlsx'

    if save_type == 'standard':
        save_path = window.app_config.get('savePath') if window.app_config.get('savePath') else window.base_save_path
        file_path = f'{save_path}/{file_name}'
    else:
        file_path, _ = QFileDialog.getSaveFileName(
            window,
            'Сохранить как Excel',
            '',
            'Excel Files (*.xlsx)'
        )
        if not file_path:
            return

    try:
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            window.result_data.to_excel(writer, index=False, sheet_name='Проценка товаров')

            workbook = writer.book
            worksheet = writer.sheets['Проценка товаров']

            formats = {
                'header': workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'border': 1
                }),
                'colored_header': workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'border': 1,
                    'bg_color': '#607ebc',
                    'font_color': '#faf5ee',
                    'align': 'right'
                }),
                'numeric_header': workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'border': 1,
                    'align': 'right'
                }),
                'data': workbook.add_format({
                    'text_wrap': True,
                    'valign': 'vcenter',
                    'font_size': 10,
                    'border': 1
                }),
                'numeric_data': workbook.add_format({
                    'text_wrap': True,
                    'valign': 'vcenter',
                    'font_size': 10,
                    'border': 1,
                    'align': 'right'
                }),
                'missing_data': workbook.add_format({
                    'text_wrap': True,
                    'valign': 'vcenter',
                    'font_size': 10,
                    'border': 1,
                    'font_color': 'red',
                    'bold': True
                })
            }

            colored_columns = [
                'Мин НАЛИЧИЕ', 'Сред НАЛИЧИЕ', 'Макс НАЛИЧИЕ',
                'Мин ПОД ЗАКАЗ', 'Сред ПОД ЗАКАЗ', 'Макс ПОД ЗАКАЗ'
            ]
            numeric_patterns = [
                'Цена магазина',
                'Кол-во магазина',
                'Кол-во дней доставки магазина'
            ]

            for col_num, column_name in enumerate(window.result_data.columns):
                is_colored = column_name in colored_columns
                is_numeric = any(pattern in column_name for pattern in numeric_patterns)
                fmt = formats['numeric_header'] if is_numeric else formats['colored_header'] if is_colored else formats['header']
                worksheet.write(0, col_num, column_name, fmt)

            for row in range(1, len(window.result_data) + 1):
                for col in range(len(window.result_data.columns)):
                    cell_value = window.result_data.iloc[row - 1, col]
                    col_name = window.result_data.columns[col]

                    is_numeric = (col_name in colored_columns or
                                  any(pattern in col_name for pattern in numeric_patterns))

                    if str(cell_value).strip() in ['Данные отсутствуют', 'Больше данных нет']:
                        fmt = formats['missing_data']
                    elif is_numeric:
                        fmt = formats['numeric_data']
                    else:
                        fmt = formats['data']

                    worksheet.write(row, col, cell_value, fmt)

            for i, column in enumerate(window.result_data.columns):
                try:
                    str_lengths = window.result_data[column].astype(str).fillna("").str.len()

                    max_len_data = str_lengths.max() if not str_lengths.empty else 0
                    max_len_column_name = len(str(column))

                    max_len = max(max_len_data, max_len_column_name)

                    width = min(50, (max_len + 2) * 1.1)

                    worksheet.set_column(i, i, width)
                except AttributeError:
                    logging.warning(f"Ошибка в столбце {column}: {str(e)}")
                    continue
                except Exception as e:
                    logging.warning(f"Ошибка в столбце {column}: {str(e)}")
                    continue

            worksheet.freeze_panes(1, 0)

    except PermissionError:
        QMessageBox.critical(
            window,
            'Ошибка доступа',
            'Нет прав на запись в указанную папку. Закройте файл если он открыт.'
        )
    except Exception as ex:
        logging.exception('Ошибка экспорта в Excel')
        QMessageBox.critical(
            window,
            'Ошибка экспорта',
            f'Не удалось экспортировать данные: {str(ex)}'
        )
