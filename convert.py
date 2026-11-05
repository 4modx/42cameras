"""Конвертирует Excel со списком камер фотовидеофиксации в data.js для карты.
Запуск: uv run --with openpyxl python -I convert.py <файл.xlsx>
"""
import json, re, sys
import openpyxl

CATS = [
    ("speed",      "Скорость",              "#e53935", r"превышение установленной скорости"),
    ("belt",       "Ремень",                "#8e24aa", r"ремнем безопасности"),
    ("phone",      "Телефон",               "#3949ab", r"телефоном"),
    ("oncoming",   "Встречная полоса",      "#d81b60", r"полосу, предназначенную для встречного"),
    ("redlight",   "Красный свет",          "#f4511e", r"запрещающий сигнал светофора"),
    ("stopline",   "Стоп-линия",            "#fb8c00", r"стоп-линией"),
    ("lights",     "Световые приборы",      "#fdd835", r"световыми приборами"),
    ("pedestrian", "Пешеходы",              "#43a047", r"уступить дорогу пешеходам"),
    ("parking",    "Остановка/стоянка",     "#00897b", r"нарушение правил остановки или стоянки"),
    ("parkhwy",    "Остановка на автодороге","#00acc1", r"остановки на дорогах для автомобилей"),
    ("signpark",   "Знаки: стоянка запрещена","#1e88e5", r"запрещающими остановку или стоянку"),
    ("trucks",     "Грузовым запрещено",    "#6d4c41", r"движение грузовых"),
    ("signs",      "Знаки/разметка",        "#546e7a", r"предписанных дорожными знаками или разметкой проезжей части дороги$"),
    ("shoulder",   "Обочина",               "#7cb342", r"по обочинам"),
    ("turn",       "Поворот/разворот",      "#c0ca33", r"крайнее положение"),
]

def classify(text):
    ids, raw = [], []
    for part in re.split(r";\s*", str(text or "")):
        part = part.strip().rstrip(".;").strip()
        if not part:
            continue
        raw.append(part)
        for cid, _, _, rx in CATS:
            if re.search(rx, part, re.I):
                if cid not in ids:
                    ids.append(cid)
                break
        else:
            print("UNCLASSIFIED:", part, file=sys.stderr)
    return ids, raw

def num(s):
    return float(str(s).strip().replace(",", "."))

PAIR = re.compile(r"(\d{2}[.,]\d{3,})\s*,\s*(\d{2}[.,]\d{3,})")

def parse_routes(text):
    """Строка вида 'lat, lon - lat, lon, lat, lon - lat, lon; ...' -> список полилиний."""
    text = str(text or "")
    matches = list(PAIR.finditer(text))
    segments, cur = [], []
    for i, m in enumerate(matches):
        pt = [num(m.group(1)), num(m.group(2))]
        if i > 0:
            between = text[matches[i-1].end():m.start()]
            if "-" not in between:
                segments.append(cur); cur = []
        cur.append(pt)
    if cur:
        segments.append(cur)
    return [s for s in segments if len(s) >= 2]

def main(path):
    ws = openpyxl.load_workbook(path, data_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    points, routes, section = [], [], "static"
    for r in rows:
        a, lat, lon, mode, viol = (list(r) + [None]*5)[:5]
        if a == "Мобильные средства фиксации":
            section = "mobile"; continue
        if a in (None, "Адреса мест установки", "Стационарные и передвижные средства фиксации") or str(a).startswith("Маршрут движения"):
            continue
        cats, raw = classify(viol)
        if section == "static":
            try:
                la, lo = num(lat), num(lon)
            except Exception:
                print("SKIP bad coords:", a, lat, lon, file=sys.stderr); continue
            points.append({"a": str(a).strip(), "la": la, "lo": lo, "m": mode or "", "c": cats, "v": raw})
        else:
            segs = parse_routes(a)
            if not segs:
                print("SKIP bad route:", str(a)[:60], file=sys.stderr); continue
            routes.append({"s": segs, "m": mode or "", "c": cats, "v": raw})
    data = {
        "source": "26.10.2026.xlsx",
        "cats": [{"id": c, "name": n, "color": col} for c, n, col, _ in CATS],
        "points": points, "routes": routes,
    }
    with open("data.js", "w", encoding="utf-8") as f:
        f.write("window.CAMERA_DATA = " + json.dumps(data, ensure_ascii=False, separators=(",", ":")) + ";\n")
    print(f"points={len(points)} routes={len(routes)} segments={sum(len(r['s']) for r in routes)}")

if __name__ == "__main__":
    main(sys.argv[1])
