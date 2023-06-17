import base.auth as auth
from googleapiclient.errors import HttpError
import json as json


def create():
    NEW_FORM = {
        "info": {
            "title": "Form đánh giá chéo bài tập lớn",
            "document_title": "Đánh giá chéo bài tập lớn"
        }
    }
    result = auth.form_service.forms().create(body=NEW_FORM).execute()
    id = result['formId']
    with open('form\\template.json', 'r') as file:
        form_body = json.load(file)
    file.close()
    request = auth.form_service.forms().batchUpdate(formId=id, body=form_body)
    request.execute()
    return id


def permission(id, email):
    permission1 = {
        'type': 'user',
        'role': 'writer',
        'emailAddress': email
    }
    auth.drive_service.permissions().create(fileId=id, body=permission1).execute()


def update(id, group, topic):
    delete(id, 3)
    delete(id, 2)
    group_json = list()
    for i in range(len(group)):
        group_json.insert(0, {'value': group[i]})
    group_json.reverse()
    topic_json = list()
    for i in range(len(topic)):
        topic_json.insert(0, {'value': topic[i]})
    topic_json.reverse()
    form_body = {
        "requests": [
            {
                "createItem": {
                    "item": {
                        "title": "T\u00ean \u0111\u1ec1 t\u00e0i?",
                        "questionItem": {
                            "question": {
                                "choiceQuestion": {
                                    "type": "DROP_DOWN",
                                    "options": topic_json
                                }
                            }
                        }
                    },
                    "location": {"index": 2}
                }
            },
            {
                "createItem": {
                    "item": {
                        "title": "Nh\u00f3m?",
                        "questionItem": {
                            "question": {
                                "choiceQuestion": {
                                    "type": "DROP_DOWN",
                                    "options": group_json
                                }
                            }
                        }
                    },
                    "location": {"index": 3}
                }
            }
        ]
    }
    print(form_body)
    request = auth.form_service.forms().batchUpdate(formId=id, body=form_body)
    respone = request.execute()
    return respone


def delete(id, index):
    delete_info = {
        "requests": [
            {
                "deleteItem": {
                    "location": {
                        "index": index
                    }
                }
            }
        ],
    }
    request = auth.form_service.forms().batchUpdate(formId=id, body=delete_info)
    respone = request.execute()
    return respone


def deleteFile(id):
    auth.drive_service.files().delete(fileId=id).execute()


def get(id):
    request = auth.form_service.forms().get(formId=id)
    respone = request.execute()
    return respone


def getFormLink(id):
    try:
        form = get(id)
    except HttpError as err:
        msg = 'link form lỗi!'
        return ''
    return form['responderUri']


def getResultLink(id):
    try:
        get(id)
    except HttpError as err:
        print(err)
        return ''
    return 'https://docs.google.com/forms/d/' + id + '/edit#responses'
