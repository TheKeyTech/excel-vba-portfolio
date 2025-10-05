
---

## 🚀 使い方 / How to Use

1. このリポジトリをZIPでダウンロード、またはクローンします。  
   *(Download this repository as ZIP or clone it.)*
2. `ピボットテーブル作成サンプル.xlsm` を開きます。  
   *(Open `ピボットテーブル作成サンプル.xlsm`.)*
3. Excelの「開発」→「マクロのセキュリティ」で「有効にする」を選択します。  
   *(Enable macros from the Developer tab.)*
4. 「操作」シートの［実行］ボタンをクリックするか、マクロ `製造実績` を実行します。  
   *(Click the “実行” button or run the macro `製造実績`.)*
5. `data/sales_data.csv` を読み込み、  
   ピボット結果が `output/ピボット集計結果.xlsx` に出力されます。  
   *(The result file will be generated in the `output` folder.)*

---

## 🧩 機能 / Features
- CSV自動取り込み（dataフォルダ内の `sales_data.csv`）  
- ピボットテーブル自動生成（Product × Month のクロス集計）  
- 集計結果をExcel形式で自動出力（`output`フォルダに保存）  
- 画面更新・再計算の制御で高速化  
- シート削除を行わない安全設計（既存データを上書きしない）

---

## 🧠 使用技術 / Technical Highlights
| 技術要素 | 説明 |
|-----------|------|
| **VBA** | Excel標準機能のみで構築（追加ライブラリ不要） |
| **PivotCaches / PivotTable** | ピボットキャッシュの動的生成 |
| **FileSystemObject** | 出力フォルダ自動生成 |
| **相対パス運用** | `ThisWorkbook.Path` 基準でどこでも動作 |
| **UI操作** | 実行ボタン（Shape）をマクロに紐付け可能 |

---

## 📚 ライセンス / License
このプロジェクトは自由に利用・改変可能です（MITライセンス推奨）。  
商用利用・教育利用ともに問題ありません。  

This project is free to use and modify under the **MIT License**.  
Commercial and educational use are both welcome.

---

## ✉️ 作者 / Author
**TheKeyTech**  
Excel VBAを中心とした業務自動化ツール開発を行っています。  
- 🌐 [ココナラ - TheKeyTech](https://coconala.com/users/5628967)  
- 💼 [クラウドワークス - TheKey（ざっきー）](https://crowdworks.jp/public/employees/6577899)  
- 🐦 [X (Twitter)](https://x.com/TheKeyTech)

---

## 🔗 外部リンク / Related Links
- GitHub: [https://github.com/yourname/excel-vba-portfolio](https://github.com/yourname/excel-vba-portfolio)
- Portfolio ZIP (動作確認用):  
  [portfolio_portable_vba_jp.zip](https://github.com/yourname/excel-vba-portfolio/releases)

---

## 💡 メモ / Tips
- CSV列名を「Product, Month, Sales」以外にしたい場合は、  
  `CreatePivotTable` 内の `.PivotFields` 名を修正するだけで対応可能です。  
- ピボットの出力先をPDFにしたい場合は、  
  `ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF` を追加すればOKです。
