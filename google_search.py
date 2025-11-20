#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import datetime
import json
import pandas as pd
import glob

from time import sleep
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

def load_config():
    with open('config.json', 'r', encoding='utf-8') as f:
        return json.load(f)

config = load_config()

GOOGLE_API_KEY          = config['google_api_key']
CUSTOM_SEARCH_ENGINE_ID = config['custom_search_engine_id']
SPREADSHEET_ID         = config['spreadsheet_id']
SHEET_NAME             = config['sheet_name']
MAX_REQUESTS           = config['max_requests']
DATA_DIR               = config['data_dir']
CREDENTIALS_FILE       = config['google_sheets_credentials_file']

def makeDir(path):
    if not os.path.isdir(path):
        os.mkdir(path)

def get_sheets_service():
    scope = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scope)
    service = build('sheets', 'v4', credentials=creds)
    return service

def readKeywordsFromSheet(max_count=100):
    keywords = []
    row_indices = []
    count = 0

    service = get_sheets_service()
    sheet = service.spreadsheets()

    # シート全体のデータを読み込み
    # シート名に特殊文字や空白が含まれる場合に対応するため、シングルクォートで囲む
    sheet_range = f"'{SHEET_NAME}'!A:Q"
    result = sheet.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=sheet_range
    ).execute()

    values = result.get('values', [])

    for row_index, row in enumerate(values):
        # 1行目（ヘッダー）をスキップ
        if row_index == 0:
            continue

        if count >= max_count:
            break

        if len(row) >= 6:  # E列とF列が存在することを確認
            e_col = row[4].strip() if len(row) > 4 else ""  # E列（インデックス4）
            f_col = row[5].strip() if len(row) > 5 else ""  # F列（インデックス5）
            n_col = row[13].strip() if len(row) > 13 else ""  # N列（インデックス13）

            if e_col and f_col and n_col != "○":  # 両方の列に値があり、N列に○がない場合
                keyword = f"{e_col} {f_col}"
                keywords.append(keyword)
                row_indices.append(row_index + 1)  # Sheets APIは1-based
                count += 1

    return keywords, row_indices

def markRowCompleted(row_index, link=None):
    service = get_sheets_service()
    sheet = service.spreadsheets()

    # 更新する値を準備
    updates = []

    # N列に○を設定
    # シート名に特殊文字や空白が含まれる場合に対応するため、シングルクォートで囲む
    updates.append({
        'range': f"'{SHEET_NAME}'!N{row_index}",
        'values': [["○"]]
    })

    # O列にlinkを設定
    if link:
        updates.append({
            'range': f"'{SHEET_NAME}'!O{row_index}",
            'values': [[link]]
        })

    # バッチ更新実行
    if updates:
        body = {
            'valueInputOption': 'RAW',
            'data': updates
        }
        sheet.values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body=body
        ).execute()

def getSearchResponse(keyword):
    today = datetime.datetime.today().strftime("%Y%m%d")
    timestamp = datetime.datetime.today().strftime("%Y/%m/%d %H:%M:%S")

    makeDir(DATA_DIR)

    service = build("customsearch", "v1", developerKey=GOOGLE_API_KEY)

    try:
        sleep(1)
        response = service.cse().list(
            q=keyword,
            cx=CUSTOM_SEARCH_ENGINE_ID,
            lr='lang_ja',
            num=5
        ).execute()
    except Exception as e:
        print(f"エラー: {e}")
        return False, None, None

    # レスポンスをjson形式で保存
    save_response_dir = os.path.join(DATA_DIR, 'response')
    makeDir(save_response_dir)
    out = {'snapshot_ymd': today, 'snapshot_timestamp': timestamp, 'keyword': keyword, 'response': response}
    jsonstr = json.dumps(out, ensure_ascii=False)
    # キーワードをファイル名に安全な形で含める
    safe_keyword = keyword.replace(' ', '_').replace('/', '_').replace('\\', '_')[:50]
    filename = f'response_{today}_{safe_keyword}.json'
    with open(os.path.join(save_response_dir, filename), mode='w') as response_file:
        response_file.write(jsonstr)

    # 5件の検索結果のlinkを取得（コンマ区切り）
    links = []
    if 'items' in response:
        for item in response['items']:
            links.append(item['link'])
    all_links = ','.join(links)

    return True, out, all_links

def makeSearchResults():
    today = datetime.datetime.today().strftime("%Y%m%d")

    response_dir = os.path.join(DATA_DIR, 'response')
    response_files = glob.glob(os.path.join(response_dir, f'response_{today}_*.json'))

    all_results = []
    cnt = 0

    for response_file_path in response_files:
        with open(response_file_path, 'r', encoding='utf-8') as response_file:
            response_json = response_file.read()
            response_tmp = json.loads(response_json)
            ymd = response_tmp['snapshot_ymd']
            keyword = response_tmp['keyword']
            response = response_tmp['response']

            if 'items' in response and len(response['items']) > 0:
                for i in range(len(response['items'])):
                    cnt += 1
                    display_link = response['items'][i]['displayLink']
                    title = response['items'][i]['title']
                    link = response['items'][i]['link']
                    snippet = response['items'][i]['snippet'].replace('\n', '')
                    all_results.append({
                        'ymd': ymd,
                        'keyword': keyword,
                        'no': cnt,
                        'display_link': display_link,
                        'title': title,
                        'link': link,
                        'snippet': snippet
                    })

    save_results_dir = os.path.join(DATA_DIR, 'results')
    makeDir(save_results_dir)

    if all_results:
        df_results = pd.DataFrame(all_results)
        df_results.to_csv(
            os.path.join(save_results_dir, f'results_{today}.tsv'),
            sep='\t',
            index=False,
            columns=['ymd', 'keyword', 'no', 'display_link', 'title', 'link', 'snippet']
        )
        print(f"結果を保存しました: results_{today}.tsv ({len(all_results)}件)")
    else:
        print("処理する結果がありませんでした")

if __name__ == '__main__':

    try:
        keywords, row_indices = readKeywordsFromSheet(MAX_REQUESTS)
        print(f"スプレッドシートから{len(keywords)}個のキーワードを読み込みました（上から優先、未処理のみ）")

        request_count = 0
        for i, (keyword, row_index) in enumerate(zip(keywords, row_indices), 1):

            print(f"処理中 ({i}/{len(keywords)}): {keyword}")
            success, json_data, link = getSearchResponse(keyword)
            if success:
                markRowCompleted(row_index, link)
                request_count += 1
                print(f"完了 - リクエスト数: {request_count}/{MAX_REQUESTS}")
            else:
                print(f"スキップ - リクエスト数: {request_count}/{MAX_REQUESTS}")

        print("検索完了。結果を整形中...")
        makeSearchResults()

    except FileNotFoundError:
        print(f"設定ファイルまたは認証ファイルが見つかりません")
    except Exception as e:
        print(f"エラーが発生しました: {e}")