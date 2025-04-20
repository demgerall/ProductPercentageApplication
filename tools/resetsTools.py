from PyQt6 import QtWidgets

from configs.configControl import saveParserConfig


def resetParseConfig(window: QtWidgets) -> None:
    """Сбрасывает настройки парсера к значениям по умолчанию и сохраняет конфиг.

    Возвращает все параметры парсера в исходное состояние:
    - Очищает путь к файлу Excel
    - Сбрасывает чекбоксы (доставка, наличие, гарантия, рейтинг)
    - Устанавливает значения спинбоксов по умолчанию
    - Отключает черный/белый списки
    - Сохраняет изменения в конфигурационный файл

    Args:
        window (QtWidgets.QWidget): Родительское окно для диалоговых сообщений.
            Должно быть виджетом из QtWidgets для корректного отображения QMessageBox.

    Side effects:
        - Сбрасывает search_file_path_Excel
        - Обновляет текст choosedFileLabel
        - Изменяет состояние всех связанных виджетов
        - Сохраняет изменения через saveParserConfig()
        - Обновляет statusLabel
    """
    window.search_file_path_Excel = ''
    window.choosedFileLabel.setText('Файл не выбран')

    window.deliveryDateCheckBox.setChecked(False)
    window.deliveryDateSpinBox.setValue(1)
    window.instockCheckBox.setChecked(False)
    window.guaranteeCheckBox.setChecked(False)
    window.rateCheckBox.setChecked(False)
    window.rateSpinBox.setValue(1)
    window.blackListCheckBox.setChecked(False)
    window.whiteListCheckBox.setChecked(False)

    saveParserConfig(window)
    window.statusLabel.setText('Настройки парсера сброшены к значениям по умолчанию')


def resetStandardSavePath(window: QtWidgets) -> None:
    """Сбрасывает путь сохранения файлов на стандартное значение.

    Устанавливает:
    - Путь сохранения в конфиге приложения в пустую строку
    - Placeholder поля ввода на базовый путь (рабочий стол пользователя)

    Args:
        window (QtWidgets.QWidget): Родительское окно для диалоговых сообщений.
            Должно быть виджетом из QtWidgets для корректного отображения QMessageBox.

    Note:
        Не сохраняет конфиг автоматически - требуется вызов saveAppConfig()
    """
    window.app_config['savePath'] = ''
    window.standardSavePathInput.setPlaceholderText(window.base_save_path)
    window.statusLabel.setText('Путь сохранения сброшен на рабочий стол')
