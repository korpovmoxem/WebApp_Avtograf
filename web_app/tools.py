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


def get_folder_id() -> str:
    """
    Считывает ID папки для загрузки отчета
    """

    with open('auth.yml') as file:
        file_data = yaml.safe_load(file)
        return str(file_data['bitrix_folder_id'])


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


def get_user_folder_id(user_id: str) -> str:
    folder_name = 'Отчет_по_активностям'
    b = fast_bitrix_client()
    storage_info = b.get_all('disk.storage.getlist', {
        'filter': {
            'ENTITY_TYPE': 'user',
            'ENTITY_ID': user_id,
        }
    })
    if not storage_info:
        return get_folder_id()

    storage_id = storage_info[0]['ID']
    storage_folders = b.get_all('disk.storage.getchildren', {
        'id': storage_id
    })
    report_folder = list(filter(lambda x: x['NAME'] == folder_name, storage_folders))
    if report_folder:
        return report_folder[0]['ID']

    new_folder = b.call('disk.storage.addfolder', {
        'id': storage_id,
        'data': {
            'NAME': folder_name
        }
    })
    return new_folder['ID']
