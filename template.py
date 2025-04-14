# Инициализация дизайна
# pyuic6 C:/Users/demge/PycharmProjects/ProductPercentageApplication/ProductPercentageApplication.ui -o C:/Users/demge/PycharmProjects/ProductPercentageApplication/ProductPercentageApplicationDesign.py

# Компилятор exe
# pyinstaller -F -w -i "C:/Users/demge/PycharmProjects/ProductPercentageApplication/franz.ico" app.py

import sys
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QCheckBox, QSpinBox, QProgressBar, QFileDialog,
    QTabWidget, QTableWidget, QTableWidgetItem, QMessageBox, QHeaderView
)
from PyQt6.QtCore import Qt


class PriceCheckerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Система проценки товаров")
        self.setFixedSize(800, 600)

        # Основной виджет и компоновка
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        # Вкладки
        self.tabs = QTabWidget()
        self.layout.addWidget(self.tabs)

        # Вкладка "Парсинг"
        self.setup_parsing_tab()

        # Вкладка "Настройки брендов"
        self.setup_brand_mapping_tab()

        # Вкладка "Списки поставщиков"
        self.setup_suppliers_tab()

        # Прогресс-бар
        self.progress_bar = QProgressBar()
        self.progress_bar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.layout.addWidget(self.progress_bar)

        # Кнопка запуска
        self.start_button = QPushButton("Запустить парсинг")
        self.start_button.clicked.connect(self.start_parsing)
        self.layout.addWidget(self.start_button)

        # Загрузка тестовых данных
        self.load_test_data()

    def setup_parsing_tab(self):
        """Вкладка с настройками парсинга"""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Загрузка Excel
        self.file_label = QLabel("Файл не выбран")
        self.browse_button = QPushButton("Выбрать файл Excel")
        self.browse_button.clicked.connect(self.load_excel)

        file_layout = QHBoxLayout()
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(self.browse_button)
        layout.addLayout(file_layout)

        # Фильтры
        self.delivery_days_check = QCheckBox("Ограничить срок доставки (дней):")
        self.delivery_days_input = QSpinBox()
        self.delivery_days_input.setRange(1, 30)

        self.in_stock_check = QCheckBox("Только в наличии")
        self.in_stock_check.setChecked(True)

        self.guarantee_check = QCheckBox("Только с гарантией наличия")

        self.rating_check = QCheckBox("Рейтинг магазина выше:")
        self.rating_input = QSpinBox()
        self.rating_input.setRange(1, 5)

        filters_layout = QVBoxLayout()
        delivery_layout = QHBoxLayout()
        delivery_layout.addWidget(self.delivery_days_check)
        delivery_layout.addWidget(self.delivery_days_input)
        filters_layout.addLayout(delivery_layout)
        filters_layout.addWidget(self.in_stock_check)
        filters_layout.addWidget(self.guarantee_check)

        rating_layout = QHBoxLayout()
        rating_layout.addWidget(self.rating_check)
        rating_layout.addWidget(self.rating_input)
        filters_layout.addLayout(rating_layout)

        layout.addLayout(filters_layout)

        self.tabs.addTab(tab, "Парсинг")

    def setup_brand_mapping_tab(self):
        """Вкладка для замены названий брендов"""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Таблица замены брендов
        self.brand_table = QTableWidget()
        self.brand_table.setColumnCount(2)
        self.brand_table.setHorizontalHeaderLabels(["Бренд в запросе", "Бренд на сайте"])
        self.brand_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        # Кнопки управления
        button_layout = QHBoxLayout()
        self.add_brand_button = QPushButton("Добавить замену")
        self.add_brand_button.clicked.connect(self.add_brand_mapping)

        self.remove_brand_button = QPushButton("Удалить выбранное")
        self.remove_brand_button.clicked.connect(self.remove_brand_mapping)

        button_layout.addWidget(self.add_brand_button)
        button_layout.addWidget(self.remove_brand_button)

        layout.addWidget(self.brand_table)
        layout.addLayout(button_layout)

        self.tabs.addTab(tab, "Настройки брендов")

    def setup_suppliers_tab(self):
        """Вкладка для управления списками поставщиков"""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        self.suppliers_table = QTableWidget()
        self.suppliers_table.setColumnCount(2)
        self.suppliers_table.setHorizontalHeaderLabels(["Бренд", "Поставщик"])

        self.load_whitelist_button = QPushButton("Загрузить белый список")
        self.load_blacklist_button = QPushButton("Загрузить черный список")

        buttons_layout = QHBoxLayout()
        buttons_layout.addWidget(self.load_whitelist_button)
        buttons_layout.addWidget(self.load_blacklist_button)

        layout.addWidget(self.suppliers_table)
        layout.addLayout(buttons_layout)

        self.tabs.addTab(tab, "Списки поставщиков")

    def load_excel(self):
        """Загрузка Excel-файла с артикулами"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл Excel", "", "Excel Files (*.xlsx)"
        )
        if file_path:
            self.file_label.setText(file_path)

    def add_brand_mapping(self):
        """Добавление замены бренда"""
        row = self.brand_table.rowCount()
        self.brand_table.insertRow(row)
        self.brand_table.setItem(row, 0, QTableWidgetItem(""))
        self.brand_table.setItem(row, 1, QTableWidgetItem(""))

    def remove_brand_mapping(self):
        """Удаление выбранных строк из таблицы брендов"""
        selected_rows = sorted(set(item.row() for item in self.brand_table.selectedItems()), reverse=True)

        if not selected_rows:
            QMessageBox.warning(self, "Ошибка", "Выберите строки для удаления!")
            return

        reply = QMessageBox.question(
            self,
            "Подтверждение",
            f"Удалить выбранные строки ({len(selected_rows)})?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            for row in selected_rows:
                self.brand_table.removeRow(row)

    def start_parsing(self):
        """Запуск парсинга (заглушка)"""
        if not self.file_label.text() or self.file_label.text() == "Файл не выбран":
            QMessageBox.warning(self, "Ошибка", "Выберите файл Excel!")
            return

        self.progress_bar.setValue(0)

        # Имитация работы
        for i in range(1, 101):
            self.progress_bar.setValue(i)
            QApplication.processEvents()

    def load_test_data(self):
        """Демо-данные для тестирования интерфейса"""
        self.brand_table.setRowCount(3)
        self.brand_table.setItem(0, 0, QTableWidgetItem("General Motors"))
        self.brand_table.setItem(0, 1, QTableWidgetItem("GM"))
        self.brand_table.setItem(1, 0, QTableWidgetItem("Land Lover"))
        self.brand_table.setItem(1, 1, QTableWidgetItem("LandRover"))
        self.brand_table.setItem(2, 0, QTableWidgetItem("Hyundai-Kia"))
        self.brand_table.setItem(2, 1, QTableWidgetItem("Hyundai/Kia/Mobis"))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PriceCheckerApp()
    window.show()
    sys.exit(app.exec())