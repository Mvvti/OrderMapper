import re
from datetime import datetime
from typing import Any

import pdfplumber


_ORDER_RE = re.compile(r"ZAM[ÓO]WIENIE\s+ZAKUPU\s+NR\s+([A-Z0-9\-/]+)", re.IGNORECASE)
_DATE_RE = re.compile(r"Data:\s*(\d{2}\.\d{2}\.\d{4})", re.IGNORECASE)
_TRADE_CONTACT_RE = re.compile(r"Sprawy\s+handlowe:\s*(.+)", re.IGNORECASE)
_PHONE_RE = re.compile(r"Phone:\s*([+\d\s\-()/]{6,})", re.IGNORECASE)
_POSTAL_CITY_RE = re.compile(r"(\d{2}-\d{3})\s+([^\d]+?)(?=\s+\d{2}-\d{3}|$)")
_ITEM_LINE_RE = re.compile(
    r"^\s*(\d+)\s+(.+?)\s+([0-9][0-9.,]*)\s+([A-Za-z]+)\s+([0-9][0-9.,]*)\s+[0-9][0-9\s.,]*\s+[0-9][0-9\s.,]*\s*$"
)
_ITEM_BOUNDARY_RE = re.compile(
    r"^\s*\d+\s+.+?\s+[0-9][0-9\s.,]*\s+[A-Za-z]+\s+[0-9][0-9\s.,]*\s+[0-9][0-9\s.,]*\s+[0-9][0-9\s.,]*\s*$"
)


def parse_pdf_order(pdf_path: str) -> dict[str, Any]:
    """Parse PDF order PDF and return normalized dictionary for mapping."""
    result: dict[str, Any] = {
        "zamowienie_nr": "",
        "data_utworzenia": "",
        "obiekt_adres_ulica": "",
        "obiekt_adres_miasto": "",
        "obiekt_adres_kod_pocztowy": "",
        "obiekt_id": "",
        "zamawiajacy_imie_nazwisko": "",
        "zamawiajacy_telefon": "",
        "pozycje": [],
    }

    try:
        with pdfplumber.open(pdf_path) as pdf:
            page_texts = [page.extract_text() or "" for page in pdf.pages]
    except Exception:
        return result

    full_text = "\n".join(page_texts)
    lines = [line.strip() for line in full_text.splitlines() if line.strip()]

    result["zamowienie_nr"] = _extract_order_number(full_text)
    result["data_utworzenia"] = _extract_iso_date(full_text)
    result["zamawiajacy_imie_nazwisko"] = _extract_contact_name(full_text)
    result["zamawiajacy_telefon"] = _extract_phone(full_text)

    street, city, postal = _extract_delivery_address(lines)
    result["obiekt_adres_ulica"] = street
    result["obiekt_adres_miasto"] = city
    result["obiekt_adres_kod_pocztowy"] = postal

    result["obiekt_id"] = _extract_object_id(lines)
    result["pozycje"] = _extract_items(lines)

    return result


def _extract_order_number(text: str) -> str:
    match = _ORDER_RE.search(text)
    return match.group(1).strip() if match else ""


def _extract_iso_date(text: str) -> str:
    match = _DATE_RE.search(text)
    if not match:
        return ""

    raw_date = match.group(1)
    try:
        parsed = datetime.strptime(raw_date, "%d.%m.%Y")
    except ValueError:
        return ""

    return parsed.strftime("%Y-%m-%dT00:00:00Z")


def _extract_contact_name(text: str) -> str:
    match = _TRADE_CONTACT_RE.search(text)
    if not match:
        return ""

    words = re.findall(r"[A-Za-zÀ-ÖØ-öø-ÿĄĆĘŁŃÓŚŹŻąćęłńóśźż-]+", match.group(1))
    if len(words) >= 2:
        return f"{words[0]} {words[1]}"
    if len(words) == 1:
        return words[0]
    return ""


def _extract_phone(text: str) -> str:
    match = _PHONE_RE.search(text)
    if not match:
        return ""

    return " ".join(match.group(1).strip().split())


def _extract_delivery_address(lines: list[str]) -> tuple[str, str, str]:
    street = ""
    city = ""
    postal = ""
    street_index = -1

    for idx, line in enumerate(lines):
        if not street and "ul." in line:
            street_match = re.search(r"ul\.\s*([^\d]*\d+[A-Za-z]?)", line, re.IGNORECASE)
            if street_match:
                street = street_match.group(1).strip(" ,")
                street_index = idx
                break

    if street_index >= 0:
        search_lines = lines[street_index:street_index + 6]
    else:
        search_lines = lines

    for line in search_lines:
        postal_city_match = _POSTAL_CITY_RE.search(line)
        if postal_city_match:
            postal = postal_city_match.group(1).strip()
            city = postal_city_match.group(2).strip(" ,")
            break

    return street, city, postal


def _extract_object_id(lines: list[str]) -> str:
    for line in lines:
        upper = line.upper()
        if upper.startswith("CD-") or "BRAMA" in upper or "GATE" in upper:
            return line.strip()
    return ""


def _extract_items(lines: list[str]) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    in_table = False
    i = 0

    while i < len(lines):
        line = lines[i]

        if "Lp." in line and "Kod odbiorcy" in line and "Cena jedn." in line:
            in_table = True
            i += 1
            continue

        if not in_table:
            i += 1
            continue

        if line in {"netto", "Nazwa pozycji"}:
            i += 1
            continue

        if line.startswith("Wartość netto") or line.startswith("Opracował"):
            in_table = False
            i += 1
            continue

        match = _ITEM_LINE_RE.match(line)
        if not match:
            i += 1
            continue

        lp = _to_int(match.group(1))
        left_part = match.group(2).strip()
        ilosc = _to_float(match.group(3))
        jednostka = match.group(4).strip()
        cena_netto = _to_float(match.group(5))

        kod_odbiorcy = _extract_receiver_code(left_part)
        nazwa, consumed = _consume_item_name(lines, i + 1)

        items.append(
            {
                "lp": lp,
                "kod_odbiorcy": kod_odbiorcy,
                "nazwa": nazwa,
                "ilosc": ilosc,
                "jednostka": jednostka,
                "cena_netto": cena_netto,
            }
        )

        i = i + 1 + consumed

    return items


def _extract_receiver_code(left_part: str) -> str:
    parts = left_part.split()
    return parts[-1] if parts else ""


def _consume_item_name(lines: list[str], start_index: int) -> tuple[str, int]:
    collected: list[str] = []
    i = start_index

    while i < len(lines):
        current = lines[i].strip()
        upper = current.upper()

        if not current:
            break
        if _ITEM_BOUNDARY_RE.match(current):
            break
        if current.startswith("Data dostawy"):
            break
        if current.startswith("Wartość netto"):
            break
        if current.startswith("Opracował"):
            break
        if current.startswith("Lp. "):
            break
        if current == "netto" or current == "Nazwa pozycji":
            break
        if "CERAMICS" in upper or upper.startswith("NIP:") or upper.startswith("PAGE "):
            break

        if not re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", current):
            collected.append(current)

        i += 1

    name = " ".join(collected).strip()
    consumed = i - start_index
    return name, consumed


def _to_float(value: str) -> float:
    compact = value.replace(" ", "")
    if "," in compact:
        normalized = compact.replace(".", "").replace(",", ".")
    else:
        normalized = compact
    try:
        return float(normalized)
    except ValueError:
        return 0.0


def _to_int(value: str) -> int:
    try:
        return int(value)
    except ValueError:
        return 0
