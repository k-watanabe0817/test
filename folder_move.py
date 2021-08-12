import pathlib
import shutil
import os
import csv

#CSVファイルのオープン（移動前、移動後ディレクトリの読み込み）
path = pathlib.Path()
for pass_obj in path.iterdir():
    if pass_obj.match("*.csv"):
        with open(pass_obj, encoding = "utf-8_sig") as f:
            reader = csv.reader(f)
            for read in reader:
                #   各pathの変数定義
                path1 = read[0]
                path2 = read[1]
                #　　指定したディレクトリへファイルを移動
                new_path = shutil.copytree(path1, path2)