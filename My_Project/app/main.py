from flask import Flask, render_template, request, redirect, url_for, session
from flask_debugtoolbar import DebugToolbarExtension
import re
import os
from urllib.parse import unquote_plus
import datetime
import json
import os.path
from google.oauth2 import service_account  # type: ignore
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# 必要なスコープ
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/drive",
]

# 環境変数からキーファイルの内容を読み込む
key_file_contents = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
credentials_info = json.loads(key_file_contents)

# 読み込んだ情報から認証情報を生成
credentials = service_account.Credentials.from_service_account_info(
    credentials_info, scopes=SCOPES
)

app = Flask(__name__, static_folder=".", static_url_path="")
app.debug = True
app.secret_key = "eihjfoq384yijf8ouawfjo"
toolbar = DebugToolbarExtension(app)

def extract_id(url):
    pattern = r"/spreadsheets/d/([a-zA-Z0-9-_]+)"
    match = re.search(pattern, url)

    if match:
        return match.group(1)
    else:
        return None


def get_sheets_data(id):
    # 対象となるスプレッドシートのIDと読み取り範囲
    SHEET_ID = id
    SHEET_NAME = "フォームの回答 1"

    try:
        service_sheets = build("sheets", "v4", credentials=credentials)

        spreadsheet = (
            service_sheets.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
        )
        title = spreadsheet["properties"]["title"]

        # Call the Sheets API
        result = (
            service_sheets.spreadsheets()
            .values()
            .get(spreadsheetId=SHEET_ID, range=SHEET_NAME)
            .execute()
        )
        values = result.get("values", [])

        return values, title

    except HttpError as err:
        print(err)
        return None


@app.route("/")
def index():
    session.clear()
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():
    url = request.form["url_input"]
    id = extract_id(url)
    sheets_data, sheets_title = get_sheets_data(id)
    # 質問取り出す
    keys = sheets_data[0]
    # シートデータ取り出す
    print("::",sheets_data)
    print("::",sheets_title)
    session["sheets_data"] = sheets_data
    session["sheets_title"] = sheets_title
    session["keys"] = keys

    return redirect(
        url_for("results")
    )


@app.route("/results")
def results():
    sheets_title = session.get("sheets_title", "titleが見つかりません")
    keys = session.get("keys", "keyが見つかりません")
    
    return render_template(
        "results.html", sheets_title=sheets_title, keys=keys
    )


@app.route("/create_document", methods=["POST"])
def write_to_google_doc():
    converted_data = session.get("sheets_data", "URLが見つかりません")
    
    # POSTリクエストからJSONデータを取得
    request_data = request.get_json()
    title = request_data.get("requestData", {}).get("title")
    selected_keys = request_data.get("requestData", {}).get("selectedKeys")

    new_list = []
    for i in range(len(converted_data[0])):
        question_data = [converted_data[0][i]]
        for j in range(1, len(converted_data)):
            if i < len(converted_data[j]):
                question_data.append(converted_data[j][i])
            else:
                question_data.append("")
        new_list.append(question_data)

    count_answer = sum(isinstance(item, list) for item in converted_data) - 1

    selected_list = [new_list[i] for i in range(len(new_list)) if selected_keys[i]]
    selected_list = [[item for item in row if item != ""] for row in selected_list]

    try:
        service = build("docs", "v1", credentials=credentials)
        body = {"title": title}
        doc = service.documents().create(body=body).execute()
        document_id = doc.get("documentId")
        document_url = f"https://docs.google.com/document/d/{document_id}/edit"
        session["document_url"] = document_url
        session["document_title"] = title

        requests = []

        # ドキュメントを逆順に構築
        for line_of_text in reversed(selected_list):
            bolded_text = f"＜{line_of_text[0]}＞\n"
            line = "_________________________________________________________________________\n\n"
            result_str = ""
            filtered_list = [x for x in line_of_text if x is not None]
            for i in range(len(filtered_list[1:])):
                result_str += f"・{filtered_list[1:][i]}\n"

            # 背景色を設定
            sub_text_style = {
                "bold": False,
            }

            requests.extend(
                [
                    {
                        "insertText": {
                            "text": bolded_text,
                            "location": {
                                "index": 1,
                            },
                        },
                    },
                    {
                        "updateTextStyle": {
                            "textStyle": {
                                "bold": True,
                            },
                            "fields": "bold",
                            "range": {
                                "startIndex": 1,
                                "endIndex": len(bolded_text),
                            },
                        },
                    },
                    {
                        "insertText": {
                            "text": result_str + "\n",
                            "location": {
                                "index": len(bolded_text) + 1,
                            },
                        },
                    },
                    {
                        "updateTextStyle": {
                            "textStyle": sub_text_style,
                            "fields": "bold,italic,backgroundColor",
                            "range": {
                                "startIndex": len(bolded_text) + 1,
                                "endIndex": len(bolded_text) + len(result_str),
                            },
                        },
                    },
                ]
            )

        requests.extend(
            [
                {
                    "insertText": {
                        "text": line,
                        "location": {
                            "index": 1,
                        },
                    },
                },
            ]
        )

        sub_header_text = (
            "回答者数："
            + str(count_answer)
            + "\n"
            + "アンケート作成日："
            + str(datetime.datetime.now().date())
            + "\n"
        )
        requests.extend(
            [
                {
                    "insertText": {
                        "text": sub_header_text,
                        "location": {
                            "index": 1,
                        },
                    },
                },
                {
                    "updateTextStyle": {
                        "textStyle": {
                            "bold": False,
                            "italic": False,
                            "fontSize": {
                                "magnitude": 11,
                                "unit": "PT",
                            },
                        },
                        "fields": "bold,italic,fontSize",
                        "range": {
                            "startIndex": 1,
                            "endIndex": len(sub_header_text),
                        },
                    },
                },
            ]
        )

        requests.extend(
            [
                {
                    "insertText": {
                        "text": line,
                        "location": {
                            "index": 1,
                        },
                    },
                },
            ]
        )

        doc_header_text = title + "\n\n"
        requests.extend(
            [
                {
                    "insertText": {
                        "text": doc_header_text,
                        "location": {
                            "index": 1,
                        },
                    },
                },
                {
                    "updateTextStyle": {
                        "textStyle": {
                            "bold": False,
                            "italic": False,
                            "fontSize": {
                                "magnitude": 17,
                                "unit": "PT",
                            },
                        },
                        "fields": "bold,italic,fontSize",
                        "range": {
                            "startIndex": 1,
                            "endIndex": len(doc_header_text),
                        },
                    },
                },
                {
                    "updateParagraphStyle": {
                        "range": {
                            "startIndex": 1,
                            "endIndex": len(doc_header_text),
                        },
                        "paragraphStyle": {
                            "alignment": "CENTER",
                        },
                        "fields": "alignment",
                    },
                },
            ]
        )

        service.documents().batchUpdate(
            documentId=document_id, body={"requests": requests}
        ).execute()

        service = build("drive", "v3", credentials=credentials)
        # 全ユーザーに編集権限を付与するためのアクセス権設定
        drive_permission = {
            "type": "anyone",  # どのユーザーでもアクセス可能
            "role": "writer",  # 編集権限
        }

        # ドキュメントにアクセス権を設定
        service.permissions().create(
            fileId=document_id, body=drive_permission, fields="id"
        ).execute()
        print("ドキュメントは全ユーザーに編集権限で共有されました。")

        return redirect(url_for("end"))
    except HttpError as err:
        print(err)


@app.route("/end", methods=["GET"])
def end():
    document_url = session.get("document_url", "URLが見つかりません")
    document_title = session.get("document_title", "URLが見つかりません")
    return render_template("end.html", document_url=document_url, document_title=document_title)


if __name__ == "__main__":
    app.run(port=8000, debug=True)
