"""Jev評価用テストデータの生成スクリプト

ルールでは「曖昧」になる金額（小数点の後ろがちょうど3桁）を20件集めたもの。
拠点・本文の言語・文脈をいろいろ組み合わせてある。
正解（truth）は、本文の文脈（何の料金か）から作成者が判断した。

実行:  .venv\\Scripts\\python.exe tools\\make_jev_testset.py
出力:  tests/jev_amount_tests.xlsx
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

ROOT = Path(__file__).resolve().parent.parent
OUT_DIR = ROOT / "tests"

# (case_id, office, amount, content, truth_separator, truth_value, note)
CASES = [
    # パリ（ルールでは , を小数点と読む）
    ("J01", "Paris", "2,400",
     "Could you confirm the total for the conference room rental?",
     "thousands", 2400, "英語の本文。会議室のレンタル料"),
    ("J02", "Paris", "3,750",
     "Our team of 25 people would like to book the seminar package.",
     "thousands", 3750, "英語の本文。25人分のセミナーパッケージ"),
    ("J03", "Paris", "1,859",
     "Le prix du litre de gazole facturé me semble trop élevé.",
     "decimal", 1.859, "フランス語の本文。軽油1リットルの値段"),
    ("J04", "Paris", "0,125",
     "Le tarif par SMS envoyé est-il correct ?",
     "decimal", 0.125, "フランス語の本文。SMS 1通の料金。先頭が0"),
    ("J05", "Paris", "4,500",
     "Nous avons commandé 150 palettes, merci de vérifier la facture.",
     "thousands", 4500, "フランス語の本文だが、150パレットで4.5ユーロはありえない"),
    ("J06", "Paris", "12,350",
     "The annual maintenance contract for our three buildings.",
     "thousands", 12350, "英語の本文。3棟の年間保守契約"),
    ("J07", "Paris", "2,750",
     "Le forfait mensuel de nettoyage de nos bureaux de 400 m².",
     "thousands", 2750, "フランス語の本文だが、400m²の清掃が2.75ユーロはありえない"),
    ("J08", "Paris", "1,075",
     "Merci de vérifier le taux de change appliqué sur notre facture en dollars.",
     "decimal", 1.075, "フランス語の本文。為替レート"),
    ("J09", "Paris", "1,500",
     "パリ支店への請求書の金額を確認してください。",
     "thousands", 1500, "日本語の本文。日本式では , は桁区切り"),
    ("J10", "Paris", "2,980",
     "オンライン講座の受講料について質問があります。",
     "thousands", 2980, "日本語の本文。講座の受講料"),
    # ニューヨーク（ルールでは . を小数点と読む）
    ("J11", "New York", "1.500",
     "Bonjour, voici la facture pour notre commande de 60 bouteilles de vin.",
     "thousands", 1500, "フランス語の本文。フランス式では . は桁区切り"),
    ("J12", "New York", "3.459",
     "Why was the fuel surcharge billed at this price per gallon?",
     "decimal", 3.459, "英語の本文。ガソリン1ガロンの値段"),
    ("J13", "New York", "4.800",
     "Nous avons reçu la facture pour l'installation de 12 ordinateurs.",
     "thousands", 4800, "フランス語の本文。パソコン12台の設置"),
    ("J14", "New York", "0.875",
     "Can you confirm the exchange rate used for the euro invoice?",
     "decimal", 0.875, "英語の本文。為替レート。先頭が0"),
    ("J15", "New York", "7.200",
     "Merci de vérifier le montant du contrat annuel de maintenance.",
     "thousands", 7200, "フランス語の本文。年間保守契約"),
    ("J16", "New York", "2.375",
     "The per-mile rate charged for the truck rental seems too high.",
     "decimal", 2.375, "英語の本文。トラックレンタルの1マイルあたりの料金"),
    ("J17", "New York", "5.000",
     "Bonjour, facture pour la location du stand au salon professionnel.",
     "thousands", 5000, "フランス語の本文。展示会のブースのレンタル"),
    # 東京（ルールでは . を小数点と読む）
    ("J18", "Tokyo", "3.500",
     "Bonjour, remboursement du déjeuner pour l'équipe, merci.",
     "thousands", 3500, "フランス語の本文。チームの昼食代（円）"),
    ("J19", "Tokyo", "150.250",
     "ユーロ建ての請求書に使われた為替レートを確認してください。",
     "decimal", 150.25, "日本語の本文。1ユーロあたりの円の為替レート"),
    ("J20", "Tokyo", "12.000",
     "Merci de vérifier la facture pour la location de la salle de réunion.",
     "thousands", 12000, "フランス語の本文。会議室のレンタル（円）"),
]

HEADERS = ["case_id", "office", "amount", "content", "truth_separator", "truth_value", "note"]


def main():
    OUT_DIR.mkdir(exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Cases"
    ws.append(HEADERS)
    for case in CASES:
        ws.append(list(case))
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for col, width in zip("ABCDEFG", [8, 10, 10, 70, 16, 12, 60]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    wb.save(OUT_DIR / "jev_amount_tests.xlsx")

    decimal = sum(1 for c in CASES if c[4] == "decimal")
    print(f"{len(CASES)} 件を作成しました（thousands {len(CASES) - decimal} 件、decimal {decimal} 件）")


if __name__ == "__main__":
    main()
