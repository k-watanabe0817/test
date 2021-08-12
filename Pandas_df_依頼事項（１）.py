import pathlib
import pandas as pd
import re

print("読み込み開始行は？？")
start_row = int(input('>>: '))

path = pathlib.Path()

for pass_obj in path.iterdir():
    if pass_obj.match("*.xlsx"):
        #DataFrameとして１つ目のsheetを読込
        input_book = pd.ExcelFile(pass_obj) 
        input_sheet_df = input_book.parse(input_book.sheet_names[0], skiprows = start_row - 1)

        #国交省作業用番号から管理者コードを抽出
        input_sheet_df["管理者コード"] = input_sheet_df["作業用番号"][1].split('-')[1]
        
        #小規模附属物の場合、ファイル名先頭に種別を追加
        input_sheet_df.to_csv(pathlib.PurePath(pass_obj).stem + '.csv', encoding = "utf-8_sig", header=True, index=False) 