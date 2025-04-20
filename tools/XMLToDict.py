import json
import requests

import xml.etree.ElementTree as ET

from typing import Any


def parseXMLResponseToDict(response: requests.Response) -> dict[str, Any]:
    """Парсит XML-ответ от API и извлекает JSON данные, преобразуя их в словарь.

    Ожидает, что ответ сервера содержит валидный XML, в теле которого находится
    JSON-строка. Извлекает JSON из текстового содержимого XML и преобразует его
    в словарь Python.

    Args:
        response (requests.Response): Ответ от сервера в XML формате, где:
            - response.text должен содержать валидный XML
            - Корневой элемент XML должен содержать JSON строку в text-атрибуте

    Returns:
        dict[str, Any]: Словарь с данными, полученными из JSON:
            - Ключи: строки (как в исходном JSON)
            - Значения: соответствующие JSON-значения (dict, list, str, int, float, bool, None)

    Raises:
        ValueError: Если возникает одна из ошибок:
            - XML не может быть распарсен
            - JSON строка невалидна
            - Корневой элемент XML не содержит текста
        TypeError: Если передан не requests.Response объект

    Examples:
        >>> # Пример с валидным ответом
        >>> response = requests.Response()
        >>> response._content = b'<root>{"key": "value"}</root>'
        >>> parseXMLResponseToDict(response)
        {'key': 'value'}

        >>> # Пример с ошибкой в JSON
        >>> response._content = b'<root>{"key": "value"</root>'
        >>> parseXMLResponseToDict(response)
        ValueError: Не удалось распарсить JSON из XML: ...
    """
    if not isinstance(response, requests.Response):
        raise TypeError(f'Ожидается requests.Response, получен {type(response).__name__}')

    try:
        root = ET.fromstring(response.text)
    except ET.ParseError as ex:
        raise ValueError(f'Не удалось распарсить XML: {ex}') from ex

    if not root.text:
        raise ValueError('Корневой элемент XML не содержит текста')

    json_str = root.text.strip()

    try:
        return json.loads(json_str)
    except json.JSONDecodeError as ex:
        raise ValueError(f'Не удалось распарсить JSON из XML: {ex}') from ex
