# 多言語Excelデータ自動化ツール

日本語・フランス語・英語が混在した問い合わせExcelを、Pythonで整形・検証・統合・分類し、レポート化する。

## フォルダ構成

```
Automation/
├── input/            処理するExcel（サンプル3拠点分）
├── output/           ツールが作るExcel（マスター、Issues、サマリー）
├── archive/          処理済みファイルの保管場所
├── answer_key/       答え合わせ用の正解表（ツールには読ませない）
├── src/              自分で書くスクリプトの置き場所
├── tools/            サンプルデータの生成スクリプト
├── .venv/            Pythonの仮想環境
├── requirements.txt  使っているライブラリとバージョン
└── .env.example      APIキーのひな形
```

## 環境の使い方（PowerShell）

仮想環境を有効にする（毎回、作業を始めるとき）:

```
.\.venv\Scripts\Activate.ps1
```

「スクリプトの実行が無効」というエラーが出たら、有効にせずに直接実行してもよい:

```
.\.venv\Scripts\python.exe src\main.py
```

サンプルデータを作り直す:

```
.\.venv\Scripts\python.exe tools\make_sample_data.py
```

## 入っているライブラリ

| ライブラリ | 用途 |
|---|---|
| pandas | Excelの読み込み・整理・出力 |
| openpyxl | 書式・条件付き書式・グラフ・シート作成 |
| dateparser | 日本語・フランス語・英語の日付の解釈 |
| langdetect | 言語判定 |
| python-dotenv | `.env` からAPIキーを読み込む |

### 注意：numpyのバージョン

このパソコンでは、最新のnumpy（2.5系）がWindowsのアプリ制御にブロックされた。
そのため numpy 2.3.5 を使っている。`requirements.txt` のバージョンは上げないこと。

### Jevを使うとき

1. 公式サイト（typesafe.ai）でAPIキーを取得する。
2. `.env.example` をコピーして `.env` を作り、キーを書き込む。
3. Python用の公式ライブラリは、公式ドキュメントに書かれているパッケージ名で入れる。
   よく似た名前の非公式パッケージに注意。

## サンプルデータ

`input/` に3拠点分、合計49行（東京15・パリ18・ニューヨーク16）。拠点ごとに見出しの言語と日付の書き方が違う。

| ファイル | 見出し | 日付の書き方 | その他 |
|---|---|---|---|
| inquiries_tokyo.xlsx | 日本語 | 年/月/日、和暦、全角 | 半角カナ・全角文字 |
| inquiries_paris.xlsx | フランス語 | 日/月/年、「1er août 2026」 | 電話番号の列がない |
| inquiries_newyork.xlsx | 英語 | 月/日/年、「July 14, 2026」 | |

たとえば `07/08/2026` は、パリでは8月7日、ニューヨークでは7月8日になる。

意図的に入れた問題（34種類、延べ59件）と、各行の正しい値は
`answer_key/answer_key.xlsx` にある。ツールの出力と比べて正解率を出すのに使う。
