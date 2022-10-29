"""Excelファイルから東証上場企業のコード一覧を取得します"""
import os
import sys
import openpyxl

ROW_DATA_START = 2
COLUMN_CODE = 1
COLUMN_NAME = 2
COLUMN_MARKET = 3
COLUMN_INDUSTRY33 = 5
COLUMN_INDUSTRY17 = 7
COLUMN_SIZE = 9

SKIP_MARKETS = ['ETF・ETN', 'PRO Market',
                'REIT・ベンチャーファンド・カントリーファンド・インフラファンド']  # 取得コードから省く市場・商品区分

def exist_keyword_in_list(keyword: str, words: list) -> bool:
    """リストの中にキーワードに指定した単語があるか調べます。

    Args:
        keyword (str): キーワード
        words (list): 確認対象のリスト

    Returns:
        bool:   True    : keywordに指定した文字列がリスト内に存在した
                False   : keywordに指定した文字列がリスト内に存在しなかった
    """
    for word in words:
        if keyword in word:
            return True
    return False

def load_worksheet(excel_file_path: str, sheet_name: str) -> openpyxl.Workbook.worksheets:
    """ワークシートの読み取り

    Args:
        file_name (str): Excelのファイル名

    Returns:
        openpyxl.Workbook.worksheets: ワークシートのシートのクラス
    """
    workbook = openpyxl.load_workbook(excel_file_path)
    worksheet = workbook[sheet_name]
    return worksheet


def get_topix_codes(excel_file_path: str, min_topix_code: int = 0) -> list:
    """Excelファイルから東証上場企業のコード一覧を取得します。

    Args:
        excel_file_path (str): ExcelファイルのPath
        min_topix_code (int, optional): 指定をすると、指定した番号以下コードは無視されます。 Defaults to 0。

    Returns:
        list: 東証上場企業コード一覧
    """
    topix_codes = []
    worksheet = load_worksheet(excel_file_path, "Sheet1")
    for row in worksheet.iter_rows(min_row=ROW_DATA_START):
        market_name = str(row[COLUMN_MARKET].value)
        if not exist_keyword_in_list(market_name, SKIP_MARKETS):
            # 途中でスクリプトが中断した際に指定のコードをスキップができる
            if row[COLUMN_CODE].value <= min_topix_code:
                continue
            output =[str(row[COLUMN_CODE].value),str(row[COLUMN_NAME].value),str(row[COLUMN_INDUSTRY33].value),str(row[COLUMN_INDUSTRY17].value),str(row[COLUMN_SIZE].value)]
            topix_codes.append(output)
    return topix_codes

