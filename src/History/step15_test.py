names = ["山田 太郎", "Marie Dupont", "佐藤花子", "田中", "Jean-Pierre de la Fontaine"]

for name in names:
    parts = name.split(" ")
    print(name, "=>", parts, len(parts))

import unicodedata

print("-----")

for c in "山田さくラÉlodie":
    print(c, unicodedata.name(c))

def is_japanese(text):
    for c in text:
        char_name = unicodedata.name(c, "")
        if "CJK" in char_name or "HIRAGANA" in char_name or "KATAKANA" in char_name:
            return True
    return False

print("-----")

for name in ["山田 太郎", "Marie Dupont", "小林 さくら", "YAMADA Taro", "yuki tanaka"]:
    print(name, "→", is_japanese(name))


def split_japanese_name(name):
    parts = name.split(" ")
    if len(parts) == 2:
        return parts[0], parts[1]
    return name, None

print("-----")

for name in ["山田 太郎", "佐藤花子", "田中", "小林 さくら"]:
    family, given = split_japanese_name(name)
    print(name, "→ family:", family, "/ given:", given)

def split_latin_name(name):
    parts = name.split(" ")

    titles = ["Mr.", "Mrs.", "Ms.", "Dr.", "Prof.", "M.", "Mme", "Mlle"]
    suffixes = ["Jr.", "Sr.", "II", "III"]

    if parts[0] in titles:
        parts = parts[1:]
    if parts[-1] in suffixes:
        parts = parts[:-1]
        
    if len(parts) == 1:
        return name, None

    upper_words = []
    other_words = []

    for word in parts:
        if word.isupper() and len(word) > 1:
            upper_words.append(word)
        else:
            other_words.append(word)

    if len(upper_words) > 0 and len(other_words) > 0:
        family = " ".join(upper_words).title()
        given = "".join(other_words)
        return family, given
    
    particles = ["de", "du", "des", "la", "le", "van", "von", "der", "di", "da"]
    for i in range(1, len(parts)):
        if parts[i] in particles:
            family = " ".join(parts[i:])
            given = " ".join(parts[:i])
            return family, given

    family = parts[-1]
    given = " ".join(parts[:-1])
    return family, given

print("-----")

for name in ["Marie Dupont", "Emily Johnson", "DUPONT Marie", "Jean-Pierre de la Fontaine", "Dr. Robert Miller Jr.", "YAMADA, Taro"]:
    family, given = split_latin_name(name)
    print(name, "→ family:", family, "/ given:", given)