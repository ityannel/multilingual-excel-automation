import unicodedata

samples = [
    "  鈴木\u3000一郎 ",
    "ｻｸﾗ商事",
    "０３－１２３４－５６７８",
    "１５，０００円",
    "Ｔ－０１３",
        "①②③",
    "㈱ミドリ",
    "ｶﾞｷﾞｸﾞ",
    "Société Générale",
    "㎏ ™ Ⅲ",
]

for s in samples:
    cleaned = unicodedata.normalize("NFKC", s).strip()
    print(f"[{s}] -> [{cleaned}]")