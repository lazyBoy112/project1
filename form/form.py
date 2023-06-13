import base.auth as auth


def create():
    NEW_FORM = {
        "info": {
            "title": "test form2",
            "document_title": "hello form 2"
        }
    }
    result = auth.form_service.forms().create(body=NEW_FORM).execute()
    return result['formId']


def update(id):
    form_body = {
        "includeFormInResponse": False,
        "requests": [
            {
                "createItem":
                {
                    "item": {
                        "title": "Họ và tên?",
                        "description": "VD: Nguyễn Văn A",
                        "questionItem": {
                            "question": {
                                "required": True,
                                "textQuestion": {
                                    "paragraph": False
                                }
                            }
                        }
                    },
                    "location": {
                        "index": 0
                    }
                }
            },
            {
                "createItem":
                {
                    "item": {
                        "title": "Mssv",
                        "description": "VD: 20203556",
                        "questionItem": {
                            "question": {
                                "required": True,
                                "textQuestion": {
                                    "paragraph": False
                                }
                            }
                        }
                    },
                    "location": {
                        "index": 1
                    }
                }
            },
            {
                "createItem":
                {
                    "item": {
                        "title": "Tên nhóm",
                        "questionItem": {
                            "question": {
                                "required": True,
                                "textQuestion": {
                                    "paragraph": False
                                }
                            }
                        }
                    },
                    "location": {
                        "index": 1
                    }
                }
            },
        ],
        # "writeControl":{
        #
        # }
    }
    request = auth.form_service.forms().batchUpdate(formId=id, body=form_body)
    respone = request.execute()
    return respone


def delete(id, index):
    delete_info = {
        "includeFormInResponse": False,
        "requests": [
            {
                "deleteItem":
                {
                    "location": {
                        "index": index
                    }
                }
            }
        ],
    }
    request = auth.form_service.forms().delete(formId=id, body=delete_info)
    respone = request.execute()
    return respone


def get(id):
    request = auth.form_service.forms().get(formId=id)
    respone = request.execute()
    return respone
