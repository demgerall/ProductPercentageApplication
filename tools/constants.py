class AppConstants:
    COLUMNS = {
        'SEARCH': ['Производитель', 'Артикул'],
        'LISTS': ['Бренд', 'Магазин'],
        'RESULT': [
            'Бренд', 'Артикул', 'Мин НАЛИЧИЕ', 'Сред НАЛИЧИЕ',
            'Макс НАЛИЧИЕ', 'Мин ПОД ЗАКАЗ', 'Сред ПОД ЗАКАЗ',
            'Макс ПОД ЗАКАЗ'
        ]
    }
    CONFIG_FILES = {
        'app': 'appConfig.json',
        'parser': 'parserConfig.json'
    }
    API_TIMEOUT = 10
