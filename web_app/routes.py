from datetime import datetime

from flask import request

from web_app import app
from web_app.modules.CreateActivitiesReport import create_activities_report


# Словарь функций для вызова из кастомного запроса

custom_webhooks = {
    'create_activities_report': create_activities_report,
}


def write_logs(request_info):
    with open('routes_logs.txt', 'a') as file:
        file.write(f"{datetime.now().strftime('%d.%m.%Y %H:%M:%S')} {request_info.args['job']}: {request_info.args}\n")


# Обработчик кастомных вебхуков Битрикс
@app.route('/bitrix/custom_webhook', methods=['POST', 'HEAD'])
def custom_webhook():
    write_logs(request)
    job = request.args['job']
    custom_webhooks[job](request.args)
    return 'OK'


