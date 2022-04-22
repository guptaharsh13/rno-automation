from auth import authenticate
from pprint import pprint


def createSheet(service, sheet_name):

    try:
        sheet = service.spreadsheets().create(
            body={"properties": {"title": sheet_name}}, fields='spreadsheetId').execute()
        print(
            f"Sheet has been created\nSheet name - {sheet_name}\nSheet Id - {sheet.get('spreadsheetId')}")
        return sheet.get('spreadsheetId')
    except:
        quit({"error": f"Unable to create sheet\nSheet name - {sheet_name}"})


def addSheets(service, sheet_names, sheet_id):
    try:
        responses = []
        for sheet_name in sheet_names:
            request_body = {
                'requests': [{
                    'addSheet': {
                        'properties': {
                            'title': sheet_name,
                        }
                    }
                }]
            }
            responses.append(service.spreadsheets().batchUpdate(
                spreadsheetId=sheet_id, body=request_body).execute())
        return responses
    except:
        quit({"error": f"Unable to add subsheets to sheet\nSheet Id - {sheet_id}"})


def updateColor(service, sheet_id, sub_sheet_id):
    request_body = {
        "requests": [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sub_sheet_id,
                        "startRowIndex": 2,
                        "endRowIndex": 3,
                        "startColumnIndex": 6,
                        "endColumnIndex": 7
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": 1.0,
                                "green": 0.0,
                                "blue": 0.0
                            },
                            "horizontalAlignment": "CENTER",
                            "textFormat": {
                                "foregroundColor": {
                                    "red": 1.0,
                                    "green": 1.0,
                                    "blue": 1.0
                                },
                                "fontFamily": "Nunito",
                                "fontSize": 12,
                                "bold": True
                            }
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sub_sheet_id,
                        "startRowIndex": 3,
                        "endRowIndex": 4,
                        "startColumnIndex": 6,
                        "endColumnIndex": 7
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": 1.0,
                                "green": 1.0,
                                "blue": 0.0
                            },
                            "horizontalAlignment": "CENTER",
                            "textFormat": {
                                "fontFamily": "Nunito",
                                "fontSize": 12,
                                "bold": True
                            }
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sub_sheet_id,
                        "startRowIndex": 4,
                        "endRowIndex": 5,
                        "startColumnIndex": 6,
                        "endColumnIndex": 7
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": 0.0,
                                "green": 1.0,
                                "blue": 0.0
                            },
                            "horizontalAlignment": "CENTER",
                            "textFormat": {
                                "fontFamily": "Nunito",
                                "fontSize": 12,
                                "bold": True
                            }
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
                }
            }
        ]
    }
    try:
        service.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id, body=request_body).execute()
        print(
            f"Subsheet has been updated with colors\nSheet Id - {sheet_id}\nSubsheet Id - {sub_sheet_id}")
    except:
        quit({"error": f"Unable to update color\nSheet Id - {sheet_id}\nSubsheet Id - {sub_sheet_id}"})


def addValues(service, sheet_id, sub_sheet):

    data = [
        {
            "range": f"{sub_sheet}!A1:E1",
            "values": [["Added By", "Link", "Name of Community",
                        "Join Status", "First Iteration"]]
        },
        {
            "range": f"{sub_sheet}!G3:G5",
            "values": [["Kicked Out"], ["In Progress"], ["Posted"]]
        }
    ]

    body = {
        'valueInputOption': "USER_ENTERED",
        'data': data
    }

    try:
        response = service.spreadsheets().values().batchUpdate(
            spreadsheetId=sheet_id, body=body).execute()
        print(
            f"Values have been added to subsheet\nSheet Id - {sheet_id}\nSubsheet name - {sub_sheet}")
        return response
    except:
        quit({"error": "Unable to update values\nSheet Id - {sheet_id}\nSubsheet Id - {sub_sheet_id}"})


def updateMany(service, sheet_id, sub_sheets):
    responses = []
    for sub_sheet in sub_sheets:
        responses.append(
            addValues(service=service, sheet_id=sheet_id, sub_sheet=sub_sheet[0]))
        updateColor(service=service, sheet_id=sheet_id,
                    sub_sheet_id=sub_sheet[1])
    return responses


def deleteSubsheet(service, sheet_id, sub_sheet_id):
    try:
        request_body = {
            "requests": [
                {
                    "deleteSheet": {
                        "sheetId": sub_sheet_id
                    }
                }
            ]
        }
        response = service.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id, body=request_body).execute()
        print(
            f"Subsheet has been deleted\nSheet Id - {sheet_id}\nSubsheet Id - {sub_sheet_id}")
        return response
    except:
        quit({"error": "Unable to delete subsheet\nSheet Id - {sheet_id}\nSubsheet Id - {sub_sheet_id}"})


def main():

    sub_sheets = ["LinkedIn", "Reddit", "Slack", "Facebook",
                  "Discord", "Telegram", "Instagram", "Whatsapp"]

    service = authenticate("sheets")
    sheet_id = createSheet(service=service, sheet_name="Sample Sheet")

    response = addSheets(
        service=service, sheet_names=sub_sheets, sheet_id=sheet_id)
    pprint(response)

    deleteSubsheet(service=service, sheet_id=sheet_id, sub_sheet_id=0)

    sheet_metadata = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    sub_sheet_ids = [temp.get("properties").get("sheetId")
                     for temp in sheet_metadata.get('sheets', '')]

    for count in range(len(sub_sheets)):
        sub_sheets[count] = (sub_sheets[count], sub_sheet_ids[count])

    response = updateMany(
        service=service, sheet_id=sheet_id, sub_sheets=sub_sheets)
    pprint(response)


if __name__ == "__main__":
    main()
