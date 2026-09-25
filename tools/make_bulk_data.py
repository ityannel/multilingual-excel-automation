"""大量テスト用データの生成スクリプト

東京・パリ・ニューヨークの月次ファイル（2026年1月〜9月）と、
まだ設定にない拠点（大阪・リヨン）のファイルを input/ に作る。

ファイル名はすべて inquiries_<拠点>_<年-月>.xlsx の形。
消すときは、名前に「_2026-」が入っているファイルを消せばよい。

実行:  .venv\\Scripts\\python.exe tools\\make_bulk_data.py
"""

import random
from calendar import monthrange
from datetime import date
from pathlib import Path

from openpyxl import Workbook

random.seed(2026)

ROOT = Path(__file__).resolve().parent.parent
INPUT_DIR = ROOT / "input"
TODAY = date(2026, 9, 22)
MONTHS = range(1, 10)
ROWS_PER_FILE = 30

TOKYO_HEADERS = ["問い合わせID", "受付日", "氏名", "会社名", "電話番号", "金額", "問い合わせ内容"]
PARIS_HEADERS = ["N° demande", "Date de réception", "Nom complet", "Société", "Téléphone", "Montant", "Message"]
NY_HEADERS = ["Inquiry ID", "Date Received", "Full Name", "Company", "Phone", "Amount", "Message"]

JA_FAMILY = ["佐藤", "鈴木", "高橋", "田中", "伊藤", "渡辺", "山本", "中村", "小林", "加藤", "吉田", "山田",
             "松本", "井上", "木村", "林", "清水", "山口", "森", "池田", "長谷川", "五十嵐", "佐々木"]
JA_GIVEN = ["太郎", "花子", "健", "美咲", "翔", "陽菜", "大輔", "さくら", "直樹", "優", "恵", "拓也", "結衣", "蓮", "愛"]
JA_COMPANIES = ["株式会社サクラ商事", "株式会社ミドリ", "株式会社アオバ", "株式会社ヒカリ", "有限会社ツバサ", "株式会社カエデ"]
JA_MESSAGES = [
    "請求書の金額が見積もりと違っています。確認をお願いします。",
    "注文した商品がまだ届きません。配送状況を教えてください。",
    "ログインするとエラーが表示されます。",
    "来月末で契約を解約したいです。手続きを教えてください。",
    "新しいプランの見積もりをお願いできますか。",
    "請求書の宛名を変更してください。",
    "配達日を変更できますか。",
    "アプリが頻繁に落ちます。",
]

FR_GIVEN = ["Marie", "Camille", "Léa", "Chloé", "Manon", "Julie", "Lucas", "Hugo", "Thomas", "Nicolas",
            "Antoine", "Élodie", "Hélène", "Jean-Pierre", "Sophie"]
FR_FAMILY = ["Martin", "Bernard", "Dubois", "Thomas", "Robert", "Petit", "Durand", "Leroy", "Moreau",
             "Simon", "Laurent", "Lefèvre", "Roux", "Girard", "Fournier", "de la Fontaine"]
FR_COMPANIES = ["Société Générale", "Boulangerie Martin & Fils", "Lefèvre SARL", "Moreau & Associés",
                "Bernard Transports", "Petit & Cie", "Girard Immobilier"]
FR_MESSAGES = [
    "Bonjour, la facture comporte une erreur sur le montant de la TVA.",
    "Notre commande n'est toujours pas arrivée.",
    "Impossible de me connecter à mon compte depuis la mise à jour.",
    "Je souhaite résilier mon abonnement à la fin du mois.",
    "Pourriez-vous nous envoyer un devis ?",
    "Merci de mettre à jour l'adresse de facturation.",
    "Le colis est arrivé endommagé.",
]
FR_MONTHS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre"]

EN_GIVEN = ["Emily", "Michael", "Sarah", "John", "Jessica", "Ashley", "Daniel", "Olivia", "Chris",
            "Matthew", "Amanda", "Ryan", "Robert", "Jennifer", "David"]
EN_FAMILY = ["Johnson", "Brown", "O'Connor", "Smith", "Miller", "Davis", "Wilson", "García", "Martinez",
             "Lee", "Taylor", "White", "Clark", "Anderson", "Thompson"]
EN_COMPANIES = ["Brightside Inc.", "Acme Corp", "O'Connor & Partners LLC", "Miller Health",
                "Davis & Co.", "Wilson Logistics", "Taylor Media"]
EN_MESSAGES = [
    "We were charged twice for our subscription. Please refund the duplicate charge.",
    "Our shipment has not arrived yet.",
    "The dashboard keeps timing out.",
    "Please cancel our plan effective next month.",
    "Can you give us a quote for the premium plan?",
    "Our invoice shows the wrong company address.",
    "The package was delivered to the wrong address.",
]
EN_MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September"]


def random_day(month):
    last = monthrange(2026, month)[1]
    if month == TODAY.month:
        last = TODAY.day
    return date(2026, month, random.randint(1, last))


# ---------------------------------------------------------------- 1行ずつの作り方

def tokyo_row(row_id, d):
    if random.random() < 0.7:
        date_text = d.strftime("%Y/%m/%d")
    else:
        date_text = f"{d.year}年{d.month}月{d.day}日"
    family = random.choice(JA_FAMILY)
    given = random.choice(JA_GIVEN)
    name = f"{family}{given}" if random.random() < 0.03 else f"{family} {given}"
    phone = None if random.random() < 0.05 else f"03-{random.randint(1000, 9999)}-{random.randint(1000, 9999)}"
    value = random.randint(10, 500) * 100
    if random.random() < 0.03:
        amount = f"{value // 1000}.{value % 1000:03d}"          # 曖昧（Jevに回る）
    elif random.random() < 0.5:
        amount = f"¥{value:,}"
    else:
        amount = f"{value:,}円"
    return [row_id, date_text, name, random.choice(JA_COMPANIES), phone, amount, random.choice(JA_MESSAGES)]


def paris_row(row_id, d):
    if random.random() < 0.7:
        date_text = d.strftime("%d/%m/%Y")
    else:
        date_text = f"{d.day} {FR_MONTHS[d.month - 1]} {d.year}"
    name = f"{random.choice(FR_GIVEN)} {random.choice(FR_FAMILY)}"
    if random.random() < 0.1:
        parts = name.split(" ", 1)
        name = f"{parts[1].upper()} {parts[0]}"
    phone = None if random.random() < 0.05 else f"01 {random.randint(10, 99)} {random.randint(10, 99)} {random.randint(10, 99)} {random.randint(10, 99)}"
    euros = random.randint(50, 9000)
    cents = random.randint(0, 99)
    if random.random() < 0.03:
        amount = f"{euros // 1000 + 1},{random.randint(100, 999)}"   # 曖昧（Jevに回る）
    elif random.random() < 0.5:
        amount = f"{euros:,}".replace(",", " ") + f",{cents:02d} €"
    else:
        amount = f"{euros:,}".replace(",", " ") + " €"
    return [row_id, date_text, name, random.choice(FR_COMPANIES), phone, amount, random.choice(FR_MESSAGES)]


def newyork_row(row_id, d):
    if random.random() < 0.7:
        date_text = d.strftime("%m/%d/%Y")
    else:
        date_text = f"{EN_MONTHS[d.month - 1]} {d.day}, {d.year}"
    name = f"{random.choice(EN_GIVEN)} {random.choice(EN_FAMILY)}"
    phone = None if random.random() < 0.05 else f"(212) 555-{random.randint(100, 199):04d}"
    dollars = random.randint(50, 9000)
    cents = random.randint(0, 99)
    if random.random() < 0.03:
        amount = f"{dollars // 1000 + 1}.{random.randint(100, 999)}"  # 曖昧（Jevに回る）
    elif random.random() < 0.5:
        amount = f"${dollars:,}.{cents:02d}"
    else:
        amount = f"${dollars:,}"
    return [row_id, date_text, name, random.choice(EN_COMPANIES), phone, amount, random.choice(EN_MESSAGES)]


OFFICES = [
    ("tokyo", "T", TOKYO_HEADERS, tokyo_row),
    ("paris", "P", PARIS_HEADERS, paris_row),
    ("newyork", "N", NY_HEADERS, newyork_row),
]


def write_file(filename, headers, rows):
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for row in rows:
        ws.append(row)
    wb.save(INPUT_DIR / filename)


def main():
    INPUT_DIR.mkdir(exist_ok=True)
    count_files = 0
    count_rows = 0
    all_rows = {}

    for keyword, prefix, headers, make_row in OFFICES:
        for month in MONTHS:
            rows = []
            for i in range(1, ROWS_PER_FILE + 1):
                row_id = f"{prefix}{month:02d}-{i:03d}"
                rows.append(make_row(row_id, random_day(month)))
            if month == 5:
                rows.append(list(rows[3]))              # 同じファイルの中の完全な重複
            all_rows[(keyword, month)] = rows
            count_rows += len(rows)

    # 別の月のファイルにも、同じ問い合わせが入ってしまったケース
    for keyword, _, _, _ in OFFICES:
        all_rows[(keyword, 9)].append(list(all_rows[(keyword, 8)][0]))
        all_rows[(keyword, 9)].append(list(all_rows[(keyword, 8)][1]))
        count_rows += 2

    # 知らない列が入っているファイル（パリ3月）
    for keyword, _, headers, _ in OFFICES:
        for month in MONTHS:
            rows = all_rows[(keyword, month)]
            if keyword == "paris" and month == 3:
                write_file(f"inquiries_{keyword}_2026-{month:02d}.xlsx",
                           headers + ["Commentaire interne"],
                           [row + ["à vérifier"] for row in rows])
            else:
                write_file(f"inquiries_{keyword}_2026-{month:02d}.xlsx", headers, rows)
            count_files += 1

    # まだ設定にない拠点（大阪＝日本、リヨン＝フランス）
    for month in [8, 9]:
        rows = [tokyo_row(f"O{month:02d}-{i:03d}", random_day(month)) for i in range(1, ROWS_PER_FILE + 1)]
        write_file(f"inquiries_osaka_2026-{month:02d}.xlsx", TOKYO_HEADERS, rows)
        rows = [paris_row(f"L{month:02d}-{i:03d}", random_day(month)) for i in range(1, ROWS_PER_FILE + 1)]
        write_file(f"inquiries_lyon_2026-{month:02d}.xlsx", PARIS_HEADERS, rows)
        count_files += 2
        count_rows += 2 * ROWS_PER_FILE

    print(f"{count_files} 個のファイル、合計 {count_rows} 行を作成しました。")


if __name__ == "__main__":
    main()
