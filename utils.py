import hashlib
from config import ICON_COLORS, FIELD_ALIASES

def safe_float(value):
    """Coerces a coordinate cell to a float, tolerating stray commas/whitespace."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip()
        try:
            return float(text)
        except ValueError:
            pass
        parts = [p.strip() for p in text.split(",") if p.strip()]
        for part in reversed(parts):
            try:
                return float(part)
            except ValueError:
                continue
    return None

def _row_dict(header, row):
    record = {}
    for index, value in enumerate(row):
        field = str(header[index]).strip() if index < len(header) and header[index] else f"Column {index + 1}"
        if field in record:
            field = f"{field} ({index + 1})"
        if value is not None:
            record[field] = value
    return record

def _normalise_field_name(value):
    return "".join(character.lower() for character in str(value) if character.isalnum())

def _find_field(record, field_type):
    normalized = [(_normalise_field_name(name), value) for name, value in record.items()]
    aliases = FIELD_ALIASES[field_type]

    for alias in aliases:
        for name, value in normalized:
            if name == alias and value is not None and str(value).strip():
                return value

    if field_type == "name":
        for name, value in normalized:
            if "projectname" in name and value is not None and str(value).strip():
                return value

    if field_type == "category":
        # Catches columns like "Sponsor (GPC/GTC/MEAG/DU/SAV)" whose
        # normalized form ("sponsorgpcgtcmeagdusav") won't exact-match the
        # plain "sponsor" alias above.
        for name, value in normalized:
            if "sponsor" in name and value is not None and str(value).strip():
                return value

    return None

def _source_color(source_name):
    digest = hashlib.sha256(source_name.encode("utf-8")).digest()
    return ICON_COLORS[int.from_bytes(digest[:4], "big") % len(ICON_COLORS)]

def _clean_text(value, limit=1200):
    if value is None:
        return ""
    if not isinstance(value, (str, int, float)):
        return ""
    return str(value).strip()[:limit]