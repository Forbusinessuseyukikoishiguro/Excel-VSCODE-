# Excel-VSCODE-
Excelのタブを自動生成＿VSCODEで
了解です！🌸
この「Excel にモジュール1〜Nのタブを自動で追加するスクリプト」を **note 記事風**にまとめてみました。
note 読者向けに、やさしく・分かりやすく・実用的に書いています。

---

# PythonでExcelに「モジュール1〜N」のシートを一括作成する方法

こんにちは、友季子です。
今日は **Python を使って Excel のシート（タブ）を連番で自動作成する方法** をご紹介します。

「教材用にシートをたくさん用意したい」
「毎回コピーしてタブを増やすのが面倒…」

そんなときに役立つ小技です。

---

## 1. 使うライブラリ

Excel 操作には **openpyxl** というライブラリを使います。

インストールしていない方は、以下を実行してください👇

```bash
pip install openpyxl
```

---

## 2. サンプルコード

以下のスクリプトを `tabout.py` という名前で保存してください。

```python
from openpyxl import load_workbook

# Excelファイルのパスを指定
excel_path = r"C:\Users\yukik\Downloads\0929test.xlsx"

# Excelを読み込み
wb = load_workbook(excel_path)

# 作成するシートの数
num_sheets = 10  

# 連番で「モジュール1」「モジュール2」…というシートを追加
for i in range(1, num_sheets + 1):
    sheet_name = f"モジュール{i}"
    if sheet_name not in wb.sheetnames:
        wb.create_sheet(title=sheet_name)

# 保存
wb.save(excel_path)
print("✅ モジュール1〜モジュール10 のシートを作成しました！")
```

---

## 3. 実行結果

このコードを実行すると、Excel を開いたときに👇のようなタブが並びます。

```
モジュール1 | モジュール2 | モジュール3 | ... | モジュール10
```

自動で連番付きのシートを作れるので、教材作成や業務の効率化に便利です。

---

## 4. カスタマイズ例

* **開始番号を変更する**

  ```python
  for i in range(5, 15):  # モジュール5〜モジュール14
  ```

* **ゼロ埋めしたい場合**

  ```python
  sheet_name = f"モジュール{i:03}"  
  # → モジュール001, モジュール002, ...
  ```

---

## 5. まとめ

* Python と openpyxl を使えば Excel のシートを自動で追加できる
* 名前は自由にカスタマイズ可能（「モジュール」「課題」「Step」など）
* 手作業でシートをコピーする手間を省ける

ちょっとした時短ですが、日々の作業をぐっと楽にしてくれると思います✨

---

* **Qiita風の技術寄り解説**に寄せたほうがいいですか？
* それとも **noteっぽく日常の工夫シェア感**を強めますか？
