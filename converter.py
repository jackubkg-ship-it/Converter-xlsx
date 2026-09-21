"""
Konvertor: köhnə format (S/S, Tarix, İşin təsviri, Say, Qiymət, DQ №-si, NV növü, ...)
-> yeni format (Avtomobilin Markası, Dövlət qeydiyyat nişanı, Servisa daxil olan tarix,
   Ehtiyyat hissəsinin qiyməti, Görülən işin qiyməti, Cəmi, Toplam xərc, ...).

Köhnə formatda 1 sətir = 1 iş qələmi (ehtiyat hissəsi VƏ YA iş haqqı).
Yeni formatda 1 sətir = 1 servis vizit(i) - eyni maşının eyni gündəki (və eyni iş
nömrəli) bütün qələmləri bir sətirdə birləşdirilir.
"""
import re
from datetime import datetime, date

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

AZ_FOLD = str.maketrans({
    "ə": "e", "Ə": "e", "ö": "o", "Ö": "o", "ü": "u", "Ü": "u",
    "ğ": "g", "Ğ": "g", "ş": "s", "Ş": "s", "ç": "c", "Ç": "c",
    "ı": "i", "I": "i", "İ": "i",
})


def _fold(value) -> str:
    if value is None:
        return ""
    return str(value).translate(AZ_FOLD).lower()


def _normalize_header(value) -> str:
    text = _fold(value).replace("\n", " ").strip()
    return re.sub(r"\s+", " ", text)


OLD_HEADER_MAP = {
    "s/s": "row_no",
    "tarix": "date",
    "isin tesviri": "description",
    "say": "qty",
    "olcu vahidi": "unit",
    "qiymet": "price",
    "yekun mebleg": "total",
    "cedvel sira sayi": "schedule_no",
    "sifarisci (a.s.a.)": "requester",
    "is icraisi (a.s.a.)": "confirmer",
    "isin novu": "work_type",
    "is №-si": "work_no",
    "dq №-si": "plate",
    "nv novu": "brand",
    "erazi": "area",
    "qeyd": "note",
}

NEW_HEADERS = [
    "Sıra nömrəsi", "Avtomobilin Markası", "Dövlət qeydiyyat nişanı", "Buraxılış ili",
    "Servisa daxil olan tarix", "Servisdan planlaşdırılan çıxma tarixi",
    "Servisdan faktiki çıxma tarixi", "Spidometr göstəricisi", "Sürücünün adı",
    "Avtomobili təmirə sorğunu göndərən şəxsin adı", "Dəyişdirilən ehtiyyat hissəsinin adı",
    "Ehtiyyat hissəsinin qiyməti (AZN)", "Görülən işin qiyməti (AZN)", "Təmirin qısa təsviri",
    "Sahədə təmiri təsdiqləyən şəxsin/lərin adı", "Cəmi", 0.15, 0.2, "Toplam xərc",
]

MONTHS_MAP = {
    "yanvar": 1, "january": 1, "jan": 1,
    "fevral": 2, "february": 2, "feb": 2,
    "mart": 3, "march": 3, "mar": 3,
    "aprel": 4, "april": 4, "apr": 4,
    "may": 5,
    "iyun": 6, "june": 6, "jun": 6,
    "iyul": 7, "july": 7, "jul": 7,
    "avqust": 8, "avgust": 8, "august": 8, "aug": 8,
    "sentyabr": 9, "september": 9, "sep": 9, "sept": 9,
    "oktyabr": 10, "october": 10, "oct": 10,
    "noyabr": 11, "november": 11, "nov": 11,
    "dekabr": 12, "december": 12, "dec": 12,
}


def guess_period_label(filename: str) -> str:
    """Пытается угадать 'Avqust 2026' из имени старого файла. Если не вышло - пусто."""
    name = re.sub(r"\.[^.]+$", "", filename)
    folded = _fold(name).replace("_", " ").replace("-", " ")
    year_m = re.search(r"20\d{2}", folded)
    year = year_m.group() if year_m else ""
    month_name_az = {
        1: "Yanvar", 2: "Fevral", 3: "Mart", 4: "Aprel", 5: "May", 6: "İyun",
        7: "İyul", 8: "Avqust", 9: "Sentyabr", 10: "Oktyabr", 11: "Noyabr", 12: "Dekabr",
    }
    for token in re.split(r"\s+", folded):
        t = token.strip(",.;")
        if t in MONTHS_MAP and year:
            return f"{month_name_az[MONTHS_MAP[t]]} {year}"
    return ""


def _find_header_row(ws, max_scan=6):
    for r in range(1, max_scan + 1):
        for c in range(1, ws.max_column + 1):
            if _normalize_header(ws.cell(row=r, column=c).value) == "s/s":
                return r
    return None


def list_sheets(file_path: str):
    """Возвращает [{'name':..., 'row_count':...}, ...] для листов, похожих на старый формат."""
    wb = openpyxl.load_workbook(file_path, data_only=True)
    result = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        header_row = _find_header_row(ws)
        if header_row is None:
            continue
        count = sum(
            1 for r in range(header_row + 1, ws.max_row + 1)
            if ws.cell(row=r, column=1).value not in (None, "")
        )
        result.append({"name": sheet_name.strip(), "row_count": count})
    return result


def _parse_date_old(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip().replace(",", ".").replace("/", ".")
    parts = [p for p in text.split(".") if p]
    if len(parts) != 3:
        return None
    try:
        d, m, y = (int(p) for p in parts)
        if y < 100:
            y += 2000
        return date(y, m, d)
    except (ValueError, TypeError):
        return None


def _to_float(value):
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", ".").strip())
    except ValueError:
        return None


def _fmt_num(n: float) -> str:
    """1000.0 -> '1000', 62.222 -> '62.22' - как в реальных новых файлах."""
    if n == int(n):
        return str(int(n))
    return f"{round(n, 2):g}"


def convert_sheet(file_path: str, sheet_name: str, contractor_name: str, client_name: str) -> openpyxl.Workbook:
    wb_in = openpyxl.load_workbook(file_path, data_only=True)
    ws_in = wb_in[sheet_name]
    header_row = _find_header_row(ws_in)
    if header_row is None:
        raise ValueError(f"'{sheet_name}' vərəqində köhnə format başlıqları tapılmadı")

    col_to_field = {}
    for c in range(1, ws_in.max_column + 1):
        field = OLD_HEADER_MAP.get(_normalize_header(ws_in.cell(row=header_row, column=c).value))
        if field:
            col_to_field[c] = field

    # ---------- считываем построчно и группируем в визиты ----------
    visits = {}  # key -> visit dict
    order = []
    for r in range(header_row + 1, ws_in.max_row + 1):
        raw = {field: ws_in.cell(row=r, column=c).value for c, field in col_to_field.items()}
        plate = str(raw.get("plate")).strip() if raw.get("plate") else None
        visit_date = _parse_date_old(raw.get("date"))
        if not plate or not visit_date:
            continue

        work_no = str(raw.get("work_no")).strip() if raw.get("work_no") not in (None, "") else None
        key = (plate, work_no) if work_no else (plate, visit_date.isoformat())

        if key not in visits:
            visits[key] = {
                "plate": plate, "date": visit_date, "brand": None,
                "requester": None, "confirmer": None, "notes": [],
                "item_lines": [], "parts_price_lines": [], "labor_price_lines": [],
            }
            order.append(key)
        v = visits[key]

        if raw.get("brand"):
            v["brand"] = str(raw["brand"]).strip()
        if raw.get("requester") and not v["requester"]:
            v["requester"] = str(raw["requester"]).strip()
        if raw.get("confirmer") and not v["confirmer"]:
            v["confirmer"] = str(raw["confirmer"]).strip()
        if raw.get("note"):
            note = str(raw["note"]).strip().replace("\n", " ")
            if note and note not in v["notes"]:
                v["notes"].append(note)

        description = str(raw.get("description")).strip() if raw.get("description") else None
        price = _to_float(raw.get("price")) or 0.0
        qty = _to_float(raw.get("qty")) or 1.0
        total = _to_float(raw.get("total"))
        if total is None:
            total = round(price * qty, 2)

        if description:
            v["item_lines"].append(description)

        line_text = f"{_fmt_num(price)}x{_fmt_num(qty)}={_fmt_num(total)}"
        work_type_folded = _fold(raw.get("work_type"))
        if "ehtiyat" in work_type_folded:
            v["parts_price_lines"].append(line_text)
        else:
            v["labor_price_lines"].append(line_text)

    # ---------- пишем новый файл ----------
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = sheet_name[:31] if sheet_name else "Sheet1"

    bold = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="DDDDDD")

    ws_out.cell(row=1, column=2, value=f'"{client_name}" Ltd şirkəti').font = bold
    ws_out.cell(row=2, column=2, value=f'"{contractor_name}" firmasından').font = bold
    ws_out.cell(row=3, column=2, value="Avtomobillərə göstərilən təmirin hesabatı").font = bold

    for c, header in enumerate(NEW_HEADERS, start=1):
        cell = ws_out.cell(row=4, column=c, value=header)
        cell.font = bold
        cell.fill = header_fill
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    r = 5
    for i, key in enumerate(order, start=1):
        v = visits[key]
        parts_price_text = "\n".join(v["parts_price_lines"]) or None
        labor_price_text = "\n".join(v["labor_price_lines"]) or None
        parts_total = sum(_to_float(x.split("=")[1]) or 0 for x in v["parts_price_lines"])
        labor_total = sum(_to_float(x.split("=")[1]) or 0 for x in v["labor_price_lines"])
        subtotal = round(parts_total + labor_total, 2)

        ws_out.cell(row=r, column=1, value=i)
        ws_out.cell(row=r, column=2, value=v["brand"] or "Naməlum")
        ws_out.cell(row=r, column=3, value=v["plate"])
        ws_out.cell(row=r, column=4, value=None)  # Buraxılış ili - köhnə formatda yoxdur
        ws_out.cell(row=r, column=5, value=v["date"])
        ws_out.cell(row=r, column=6, value=v["date"])
        ws_out.cell(row=r, column=7, value=v["date"])
        ws_out.cell(row=r, column=8, value=None)  # Spidometr - köhnə formatda yoxdur
        ws_out.cell(row=r, column=9, value=None)  # Sürücünün adı - köhnə formatda yoxdur
        ws_out.cell(row=r, column=10, value=v["requester"])
        ws_out.cell(row=r, column=11, value="\n".join(v["item_lines"]) or None)
        ws_out.cell(row=r, column=12, value=parts_price_text)
        ws_out.cell(row=r, column=13, value=labor_price_text)
        ws_out.cell(row=r, column=14, value="; ".join(v["notes"]) or None)
        ws_out.cell(row=r, column=15, value=v["confirmer"])
        ws_out.cell(row=r, column=16, value=subtotal)
        ws_out.cell(row=r, column=17, value=None)
        ws_out.cell(row=r, column=18, value=None)
        ws_out.cell(row=r, column=19, value=subtotal)  # надбавка köhnə formatda yoxdur -> Toplam = Cəmi

        for c in (5, 6, 7):
            ws_out.cell(row=r, column=c).number_format = "DD.MM.YYYY"
        r += 1

    widths = [8, 26, 16, 10, 14, 16, 16, 12, 16, 22, 34, 22, 22, 24, 22, 10, 8, 8, 10]
    for c, w in enumerate(widths, start=1):
        ws_out.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w

    return wb_out
