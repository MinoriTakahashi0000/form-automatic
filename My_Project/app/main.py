from flask import Flask, render_template, request, redirect, url_for
import re
import os
from urllib.parse import unquote_plus
import datetime
import webbrowser
import ast
import json
import os.path
from google.oauth2 import service_account  # type: ignore
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# 必要なスコープ
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 
          'https://www.googleapis.com/auth/documents']

# 環境変数からキーファイルの内容を読み込む
key_file_contents = os.environ.get('GOOGLE_APPLICATION_CREDENTIALS')
credentials_info = json.loads(key_file_contents)

# 読み込んだ情報から認証情報を生成
credentials = service_account.Credentials.from_service_account_info(credentials_info, scopes=SCOPES)

app = Flask(__name__, static_folder=".", static_url_path="")
app.secret_key = "hogehoge"


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
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():
    url = request.form["url_input"]
    id = extract_id(url)
    sheets_data, sheets_title = get_sheets_data(id)
    # 質問取り出す
    keys = sheets_data[0]

    return redirect(
        url_for("results", title=sheets_title, sheets_data=sheets_data, keys=keys)
    )


@app.route("/results")
def results():
    title = request.args.get("title")
    sheets_data = request.args.getlist("sheets_data")
    keys = request.args.getlist("keys")

    return render_template(
        "results.html", title=title, sheets_data=sheets_data, keys=keys
    )


@app.route("/create_document", methods=["POST"])
def write_to_google_doc():
    # POSTリクエストからJSONデータを取得
    request_data = request.get_json()

    # タイトルと選択されたキーのリストを取得
    sheets_data = request_data.get("requestData", {}).get("sheets_data")
    # 空のリストを用意
    converted_data = []
    # 各要素を処理してPythonのリストに変換
    for item in sheets_data:
        # シングルクォートをダブルクォートに変換し、JSON形式の文字列にする
        json_str = item.replace("'", '"')
        # JSON形式の文字列をPythonのリストに変換
        item_list = json.loads(json_str)
        # リストを追加
        converted_data.append(item_list)

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

        doc = service.documents().get(documentId=document_id).execute()

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

        url = "https://docs.google.com/document/d/" + document_id
        service.documents().batchUpdate(
            documentId=document_id, body={"requests": requests}
        ).execute()
        webbrowser.open(url)
        return redirect(url_for("end"))
    except HttpError as err:
        print(err)


@app.route("/end", methods=["GET"])
def end():
    return render_template("end.html")


if __name__ == "__main__":
    app.run(port=8000, debug=True)
