# 概要
エクセルの対象シートにて、指定した文字列を一括で置換するスクリプト。  
**openpyxlが必要**

# 使い方
```
python3 ExcelReplace.py "エクセルファイルの絶対パス" "シート名"
```
事前にreplace.txt に置き換えしたい文字列を以下の形式で記載しておく。

```
置換前1:置換後1
置換前2:置換後2
```
