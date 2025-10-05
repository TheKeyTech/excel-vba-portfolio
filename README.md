# 📊 Excel VBA ポートフォリオ：ピボット自動化ツール  
**Excel VBA Portfolio: Pivot Table Automation Tool**

---

## 🧭 概要 / Overview
このプロジェクトは、**CSVファイルを自動で読み込み、ピボットテーブルを作成し、結果をExcel形式で出力する**  
ポータブルなVBAツールです。  
`ThisWorkbook.Path` を基準に動作するため、**どのフォルダに配置しても実行可能**です。  
（会社固有のパスやサーバー情報は含まれていません）

This VBA project automatically imports a CSV file, creates a Pivot Table,  
and exports the result as an Excel file.  
It uses relative paths (`ThisWorkbook.Path`) so it runs **from any folder**.  
No company-specific or confidential information is included.

---

## 📁 フォルダ構成 / Folder Structure
excel-vba-portfolio/
├─ README.md
├─ ピボットテーブル作成サンプル.xlsm ← 実行用ブック（ボタン付き）
├─ ポータブル版_ピボット作成.bas ← VBAモジュール（標準モジュールにインポート可）
├─ data/
│ └─ sales_data.csv ← サンプルCSV（Product, Month, Sales）
└─ output/ ← 出力先（自動生成）
