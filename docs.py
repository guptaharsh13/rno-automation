from auth import authenticate
import traceback


def createDoc(service, doc_name):
    try:
        doc = service.documents().create(body={"title": doc_name}).execute()
    except:
        quit({"error": f"Unable to create doc\nDoc name - {doc_name}"})
    print(
        f"Doc has been created\nDoc name - {doc.get('title')}\nDoc Id - {doc.get('documentId')}")
    return doc.get("documentId")


def insertContent(service, doc_id):

    with open("doc_content.txt", "r", encoding="unicode_escape") as content:
        requests = [
            {
                'insertText': {
                    'location': {
                        'index': 1,
                    },
                    'text': content.read()
                }
            }
        ]

        try:
            response = service.documents().batchUpdate(
                documentId=doc_id, body={'requests': requests}).execute()
            print(f"Content has been added to the doc\nDoc Id - {doc_id}")
            return response
        except:
            traceback.print_exc()
            quit({"error": f"Unable to add content to doc\nDocument Id - {doc_id}"})


def basicFormat(service, doc_id):
    end_index = service.documents().get(documentId=doc_id).execute()[
        "body"]["content"][-1]["endIndex"]
    requests = [
        {
            'updateTextStyle': {
                'range': {
                    'startIndex': 1,
                    'endIndex': end_index
                },
                'textStyle': {
                    'weightedFontFamily': {
                        'fontFamily': 'Times New Roman'
                    },
                    'fontSize': {
                        'magnitude': 14,
                        'unit': 'PT'
                    }
                },
                'fields': 'weightedFontFamily,fontSize'
            }
        },
        {
            'updateParagraphStyle': {
                'range': {
                    'startIndex': 1,
                    'endIndex':  end_index
                },
                'paragraphStyle': {
                    'lineSpacing': 150
                },
                'fields': 'lineSpacing'
            }
        }
    ]

    try:
        response = service.documents().batchUpdate(
            documentId=doc_id, body={'requests': requests}).execute()
        print(f"Doc has been formatted\nDoc Id - {doc_id}")
    except:
        quit({"error": f"Doc could not be formatted\nDoc Id - {doc_id}"})
    return response


def makeHeading(service, doc_id, start_index, end_index, heading):
    requests = [
        {
            'updateParagraphStyle': {
                'range': {
                    'startIndex': start_index,
                    'endIndex':  end_index
                },
                'paragraphStyle': {
                    'namedStyleType': f'HEADING_{heading}',
                },
                'fields': 'namedStyleType'
            }
        }
    ]

    try:
        return service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()
    except:
        quit({"error": f"Heading could not be made\nDoc Id - {doc_id}"})


def makeBold(service, doc_id, start_index, end_index):
    requests = [
        {
            'updateTextStyle': {
                'range': {
                    'startIndex': start_index,
                    'endIndex': end_index
                },
                'textStyle': {
                    'bold': True
                },
                'fields': 'bold'
            }
        }
    ]

    try:
        return service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()
    except:
        quit({"error": f"Content could not be made bold\nDoc Id - {doc_id}"})


def underline(service, doc_id, start_index, end_index):
    requests = [
        {
            'updateTextStyle': {
                'range': {
                    'startIndex': start_index,
                    'endIndex': end_index
                },
                'textStyle': {
                    'underline': True
                },
                'fields': 'underline'
            }
        }
    ]

    try:
        return service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()
    except:
        quit({"error": f"Content could not be underlined\nDoc Id - {doc_id}"})


def addBullets(service, doc_id, start_index, end_index):
    requests = [
        {
            'createParagraphBullets': {
                'range': {
                    'startIndex': start_index,
                    'endIndex':  end_index
                },
                'bulletPreset': 'NUMBERED_DECIMAL_ALPHA_ROMAN',
            }
        }
    ]

    try:
        return service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()
    except:
        quit({"error": f"Bullet list could not be made\nDoc Id - {doc_id}"})


def formatDoc(service, doc_id):

    basicFormat(service=service, doc_id=doc_id)

    response = service.documents().get(documentId=doc_id).execute()

    lines = response["body"]["content"][1:]

    # end of the document
    end_index = response["body"]["content"][-1]["endIndex"]

    for line in lines:
        if line['endIndex'] - line['startIndex'] <= 1:
            continue

        print(
            f"start_index - {line['startIndex']}\nend_index - {line['endIndex']}")
        content = line["paragraph"]["elements"][0]["textRun"]["content"]
        print(f"content - {content}")

        if "Templates" in content and line['startIndex'] > 200:
            addBullets(service=service, doc_id=doc_id,
                       start_index=line['startIndex'], end_index=end_index)
            break

        if "Templates" in content or "Mail Template" in content or "Timeline" in content:
            makeBold(service=service, doc_id=doc_id,
                     start_index=line['startIndex'], end_index=line['endIndex'])

        if "Template" not in content:
            #  and not ("Time" in content or "Mails" in content)
            underline(service=service, doc_id=doc_id,
                      start_index=line['startIndex'], end_index=line['endIndex'])

    print(f"Doc has been formatted\nDoc Id - {doc_id}")


def main():
    service = authenticate("docs")
    doc_id = createDoc(service=service, doc_name="Sample Doc")
    insertContent(service=service, doc_id=doc_id)
    formatDoc(service=service, doc_id=doc_id)


if __name__ == '__main__':
    main()
