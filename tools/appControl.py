import logging

from PyQt6 import QtWidgets

from tools.configControl import saveAppConfig, saveParserConfig

from tools.tableControl import validateTable


def changePage(window: QtWidgets, index: int) -> None:
    """Осуществляет переход между страницами интерфейса с сохранением данных.

    Args:
        window (QtWidgets.QWidget): Родительское окно для диалоговых сообщений.
            Должно быть виджетом из QtWidgets для корректного отображения QMessageBox.
        index (int): Индекс целевой страницы (0-5)

    Raises:
        ValueError: Если передан недопустимый индекс страницы
        RuntimeError: При ошибках сохранения конфигурации

    Note:
        Перед переходом выполняет:
        1. Валидацию данных на текущей странице
        2. Сохранение соответствующих настроек
        3. Обновление UI при необходимости
    """
    if not 0 <= index < window.stackedWidget.count():
        raise ValueError(f'Недопустимый индекс страницы: {index}')

    try:
        current_page = window.stackedWidget.currentIndex()

        if current_page == 0:
            saveParserConfig(window)

        elif current_page == 1:
            if not validateTable(window, window.brandsTable):
                return
            saveParserConfig(window)

        elif current_page in {2, 3}:
            if not validateTable(
                    window,
                    window.blackListTable if current_page == 2 else window.whiteListTable
            ):
                return
            saveParserConfig(window)
            updateTableLabels(window, current_page)

        elif current_page == 4:
            saveAppConfig(window)

        window.stackedWidget.setCurrentIndex(index)

    except Exception as ex:
        logging.error(f'Ошибка перехода на страницу {index}: {ex}')
        raise RuntimeError(f'Не удалось перейти на страницу: {ex}') from ex


def updateTableLabels(window: QtWidgets, index: int) -> None:
    """
    Обновляет текстовые метки, отображающие количество элементов в таблицах черного/белого списков.

    В зависимости от переданного индекса обновляет соответствующую метку:
    - Индекс 2: метка для черного списка (blackListEntitiesAmountLabel)
    - Индекс 3: метка для белого списка (whiteListEntitiesAmountLabel)

    Args:
        window (QtWidgets.QWidget): Родительское окно для диалоговых сообщений.
            Должно быть виджетом из QtWidgets для корректного отображения QMessageBox.
        index (int): Индекс вкладки, определяющая какую метку обновлять:
                    * 2 - для черного списка
                    * 3 - для белого списка

    Note:
        - Формат текста метки: "(N записи(-ей))", где N - количество строк в таблице
        - Функция не производит действий для индексов, отличных от 2 или 3
        - Для получения количества строк используется метод rowCount() соответствующей таблицы
    """
    if index == 2:
        row_count = window.blackListTable.rowCount()
        window.blackListEntitiesAmountLabel.setText(
            f'({row_count} {"запись" if row_count == 1 else "записи" if 2 <= row_count <= 4 else "записей"})'
        )

    elif index == 3:
        row_count = window.whiteListTable.rowCount()
        window.whiteListEntitiesAmountLabel.setText(
            f'({row_count} {"запись" if row_count == 1 else "записи" if 2 <= row_count <= 4 else "записей"})'
        )
