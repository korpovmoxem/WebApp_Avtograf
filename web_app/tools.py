import requests
import yaml

from fast_bitrix24 import Bitrix


def bitrix_auth() -> str:
    """
    Считывает токен вебхука Б24 из файла
    """

    with open('auth.yml') as file:
        file_data = yaml.safe_load(file)
        return file_data['b24_webhook']


def fast_bitrix_client() -> Bitrix:
    """
    Создает объект класса Bitrix из библиотеки fast_bitrix24 с токеном вебхука в параметрах
    Используется для выгрузки большого массива данных
    """

    bitrix_token = bitrix_auth()
    return Bitrix(bitrix_token)


def send_bitrix_request(method: str, data=None) -> dict | list:
    """
    Отправляет запрос в Б24. Не подходит для выгрузки большого массива данных (больше 50)

    :param method: Метод запроса в Б24
    :param data: Словаь с параметрами запроса
    :return: Ответ от Б24
    """

    bitrix_token = bitrix_auth()
    request_json = requests.post(f"{bitrix_token}{method}", json=data).json()

    if 'result' in request_json:
        return request_json['result']
    print(request_json)
