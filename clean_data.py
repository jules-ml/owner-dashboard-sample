#!/usr/bin/env python3
"""Clean and validate the messy Lantern Key Pool & Air sample exports.

Reads data/raw/, writes data/clean/ tables plus a data-quality log.
All logic is rule-based (no hardcoded customer answer key).
"""
from __future__ import annotations

import json
import re
import math
from collections import defaultdict
from dataclasses import dataclass, field, asdict
from datetime import date, datetime, timedelta
from difflib import SequenceMatcher
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
RAW = ROOT / "data" / "raw"
CLEAN = ROOT / "data" / "clean"
AS_OF = date(2026, 11, 25)

CANONICAL_SERVICES = (
    "Weekly Pool Service",
    "Pool Repair",
    "HVAC Maintenance",
    "HVAC Repair",
    "Equipment Install",
)

SERVICE_RULES: list[tuple[str, str]] = [
    # Order matters: more specific first.
    ("weekly pool", "Weekly Pool Service"),
    ("pool service", "Weekly Pool Service"),
    ("pool clean", "Weekly Pool Service"),
    ("pool svc", "Weekly Pool Service"),
    ("pump / filter", "Pool Repair"),
    ("pool equip", "Pool Repair"),
    ("pool repair", "Pool Repair"),
    ("heater install", "Equipment Install"),
    ("ac replacement", "Equipment Install"),
    ("equipment install", "Equipment Install"),
    ("installation", "Equipment Install"),
    ("hvac pma", "HVAC Maintenance"),
    ("hvac maint", "HVAC Maintenance"),
    ("a/c maintenance", "HVAC Maintenance"),
    ("ac tune", "HVAC Maintenance"),
    ("tune-up", "HVAC Maintenance"),
    ("hvac-repair", "HVAC Repair"),
    ("hvac repair", "HVAC Repair"),
    ("air conditioning repair", "HVAC Repair"),
    ("a/c repair", "HVAC Repair"),
    ("ac service", "HVAC Repair"),
    # Bare "install" last.
    ("install", "Equipment Install"),
]

NICKNAMES = {
    "jim": "james",
    "jimmy": "james",
    "bob": "robert",
    "bobby": "robert",
    "mike": "michael",
    "tony": "anthony",
    "chris": "christopher",
    "tom": "thomas",
    "tommy": "thomas",
    "pat": "patrick",
    "alex": "alexander",
    "liz": "elizabeth",
    "beth": "elizabeth",
    "lauren": "lauren",
}

ENTITY_SUFFIXES = re.compile(
    r"\b(llc|inc|inc\.|incorporated|h\.?o\.?a\.?|hoa|association|assoc|assn|"
    r"condos?|club|property management|property mgmt\.?|pm)\b",
    re.I,
)

BLANK_TOKENS = {"", "n/a", "na", "none", "null", "-", "--"}
GARBAGE_TOKENS = {"tbd", "?", "??", "#ref!", "see notes", "not recorded", "00/00/0000"}

LEAD_STATUS_MAP = {
    "won": "Won",
    "lost": "Lost",
    "quoted": "Quoted",
    "contacted": "Contacted",
    "new": "New",
}

LEAD_SOURCE_MAP = {
    "google": "Google",
    "referral": "Referral",
    "yard sign": "Yard Sign",
    "nextdoor": "Nextdoor",
    "angi": "Angi",
    "repeat customer": "Repeat Customer",
    "facebook": "Facebook",
    "fb": "Facebook",
}


@dataclass
class QAEvent:
    kind: str
    message: str
    record: str = ""
    extra: dict = field(default_factory=dict)


class QualityLog:
    def __init__(self) -> None:
        self.events: list[QAEvent] = []
        self.counts: dict[str, int] = defaultdict(int)

    def add(self, kind: str, message: str, record: str = "", **extra) -> None:
        self.events.append(QAEvent(kind=kind, message=message, record=record, extra=extra))
        self.counts[kind] += 1

    def count(self, kind: str) -> int:
        return int(self.counts.get(kind, 0))


# ---------------------------------------------------------------------------
# Parsers
# ---------------------------------------------------------------------------

def is_blank(value) -> bool:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return True
    if pd.isna(value):
        return True
    return str(value).strip().lower() in BLANK_TOKENS


def parse_money(value):
    if is_blank(value):
        return None
    if isinstance(value, str) and value.strip().lower() in GARBAGE_TOKENS:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    s = str(value).strip()
    if s.strip().lower() in GARBAGE_TOKENS:
        return None
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1]
    s = re.sub(r"(?i)usd", "", s).replace("$", "").strip()
    if s.startswith("-"):
        negative = True
        s = s[1:].strip()
    # Decide whether comma is a decimal or a thousands separator.
    if "." in s and "," in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")  # 1.250,00
        else:
            s = s.replace(",", "")  # 1,250.00
    elif "," in s:
        if re.fullmatch(r"\d{1,3}(?:,\d{3})+", s):
            s = s.replace(",", "")  # 1,250
        elif re.fullmatch(r"\d+,\d{1,2}", s):
            s = s.replace(",", ".")  # 350,00
        elif re.fullmatch(r"\d{1,3}(?:,\d{3})+,\d{1,2}", s):
            whole, frac = s.rsplit(",", 1)
            s = whole.replace(",", "") + "." + frac
        else:
            s = s.replace(",", "")
    try:
        amount = float(s)
    except ValueError:
        return None
    return -amount if negative else amount


def parse_date(value, log: QualityLog | None = None, record: str = ""):
    """Parse mixed date formats. Prefer US month-first for ambiguous values."""
    if is_blank(value):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        # Excel serial (1899-12-30 system), typical 2000s range.
        if 40000 < float(value) < 60000:
            return (datetime(1899, 12, 30) + timedelta(days=int(value))).date()
        # Sometimes Excel serials arrive as floats like 46096.0
        if 40000 < float(value) < 60000:
            return (datetime(1899, 12, 30) + timedelta(days=int(value))).date()
    s = str(value).strip()
    # Pure serial written as string
    if re.fullmatch(r"\d{5}(?:\.0+)?", s):
        return (datetime(1899, 12, 30) + timedelta(days=int(float(s)))).date()

    formats = (
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%m/%d/%y",
        "%Y/%m/%d",
        "%m.%d.%Y",
        "%d-%b-%Y",
        "%b-%d-%Y",
        "%B %d, %Y",
        "%d-%B-%Y",
        "%B %d %Y",
    )
    for fmt in formats:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue

    # Day-first unambiguous (day > 12) or explicit dd/mm/yyyy when month slot > 12.
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{2,4})", s)
    if m:
        a, b, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
        if a > 12 and 1 <= b <= 12:
            try:
                return date(y, b, a)
            except ValueError:
                return None
        if b > 12 and 1 <= a <= 12:
            try:
                return date(y, a, b)
            except ValueError:
                return None
        # Ambiguous: prefer US month-first and flag.
        try:
            parsed = date(y, a, b)
            if log is not None:
                log.add("ambiguous_date", f"Ambiguous date '{s}' parsed month-first as {parsed.isoformat()}", record)
            return parsed
        except ValueError:
            try:
                return date(y, b, a)
            except ValueError:
                return None
    return None


def normalize_invoice_number(value) -> str | None:
    if is_blank(value):
        return None
    digits = re.sub(r"\D", "", str(value))
    return digits or None


def normalize_phone(value) -> str | None:
    if is_blank(value):
        return None
    digits = re.sub(r"\D", "", str(value))
    if digits.startswith("1") and len(digits) == 11:
        digits = digits[1:]
    if len(digits) == 10:
        return digits
    return None


def normalize_email(value) -> str | None:
    if is_blank(value):
        return None
    s = str(value).strip().lower()
    s = s.replace("@@", "@")
    if re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", s):
        return s
    return None


def canonicalize_service(value) -> str | None:
    if is_blank(value):
        return None
    s = str(value).strip().lower()
    s = s.replace("&", " ").replace("-", " ")
    s = re.sub(r"\s+", " ", s)
    for needle, canon in SERVICE_RULES:
        if needle in s:
            return canon
    return None


def normalize_name(value) -> str:
    if is_blank(value):
        return ""
    s = str(value).strip()
    s = s.replace("&", " and ")
    # Last, First
    if "," in s:
        parts = [p.strip() for p in s.split(",", 1)]
        if len(parts) == 2 and parts[0] and parts[1]:
            s = f"{parts[1]} {parts[0]}"
    s = s.lower()
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    tokens = [NICKNAMES.get(t, t) for t in s.split() if t]
    # Drop 1-character tokens that are leftover initials later during matching,
    # but keep them here so "m gonzalez" still has a first-name signal.
    return " ".join(tokens)


def core_name(value) -> str:
    """Name with legal-entity suffixes removed, for HOA/commercial matching."""
    if is_blank(value):
        return ""
    s = str(value).strip()
    s = ENTITY_SUFFIXES.sub(" ", s)
    return normalize_name(s)


def proper_case_name(name: str) -> str:
    small = {"and", "of", "the"}
    acronyms = {"hoa", "llc", "inc", "inc.", "pm", "h.o.a."}

    def cap_word(p: str) -> str:
        key = p.lower().strip(".,")
        if key in acronyms:
            if key in {"inc", "inc."}:
                return "Inc."
            return "H.O.A." if key == "h.o.a." else key.upper()
        if key in small:
            return key
        if "'" in p:
            return "'".join(cap_word(part) if part else "" for part in p.split("'"))
        return p[:1].upper() + p[1:].lower() if p else p

    parts = [p for p in re.split(r"(\s+)", name.strip()) if p]
    return "".join(p if not p.strip() else cap_word(p) for p in parts)


# ---------------------------------------------------------------------------
# Customer matching (union-find)
# ---------------------------------------------------------------------------

class UnionFind:
    def __init__(self, n: int) -> None:
        self.p = list(range(n))

    def find(self, x: int) -> int:
        while self.p[x] != x:
            self.p[x] = self.p[self.p[x]]
            x = self.p[x]
        return x

    def union(self, a: int, b: int) -> None:
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.p[rb] = ra


def _name_tokens(norm: str) -> list[str]:
    return [t for t in norm.split() if t]


def names_similar(a: str, b: str) -> bool:
    if not a or not b:
        return False
    if a == b:
        return True
    ratio = SequenceMatcher(None, a, b).ratio()
    if ratio >= 0.93:
        return True
    # Same tokens in different order, or a typo in one token (Gonzalez/Gonzales).
    ta, tb = _name_tokens(a), _name_tokens(b)
    if not ta or not tb:
        return False
    if set(ta) == set(tb):
        return True
    # Compare first full token + last token with edit-ish similarity.
    if ta[-1] and tb[-1] and SequenceMatcher(None, ta[-1], tb[-1]).ratio() >= 0.84:
        firsts_a = [t for t in ta[:-1] if len(t) > 1]
        firsts_b = [t for t in tb[:-1] if len(t) > 1]
        if firsts_a and firsts_b and (firsts_a[0] == firsts_b[0] or firsts_a[0][0] == firsts_b[0][0]):
            return True
        # Initial + last name: "m gonzalez" vs "maria gonzalez"
        if (not firsts_a or not firsts_b) and (ta[0][0] == tb[0][0]):
            return True
    return False


def _flip_last_first(raw: str) -> str:
    s = re.sub(r"\s+", " ", (raw or "").strip())
    if "," in s and not ENTITY_SUFFIXES.search(s):
        left, right = [p.strip() for p in s.split(",", 1)]
        if left and right:
            return f"{right} {left}"
    return s


def choose_display_name(mentions: list[dict], emails: list[str]) -> str:
    """Pick a readable canonical name from messy variants."""
    variants = [(m.get("raw_name") or "").strip() for m in mentions if (m.get("raw_name") or "").strip()]
    if not variants:
        return "Unknown"
    email_tokens = set()
    for em in emails or []:
        local = em.split("@")[0].lower()
        local = local.replace("sample", " ")
        email_tokens.update(t for t in re.split(r"[._+\-]+", local) if len(t) > 1)

    formal_firsts = set(NICKNAMES.values())

    def score(raw: str) -> tuple:
        s = _flip_last_first(raw)
        sl = s.lower()
        sc = 0.0
        if "'" in s or "'" in raw:
            sc += 5
        compact = re.sub(r"[^a-z0-9]+", "", sl)
        for tok in email_tokens:
            if tok in sl or tok in compact:
                sc += 4
        first = re.split(r"[\s,]+", s)[0].lower().strip(".")
        if first in NICKNAMES:
            sc -= 4
        if first in formal_firsts:
            sc += 2
        if re.fullmatch(r"[A-Za-z]\.?", re.split(r"[\s,]+", s)[0]):
            sc -= 5
        if re.search(r"\bHOA\b", s, re.I) and not re.search(r"\bInc\b", s, re.I):
            sc += 3
        if re.search(r"hoa", sl) and re.search(r"\binc\b", sl):
            sc -= 3
        if re.search(r"\bAssociation\b", s, re.I):
            sc += 2
        if re.search(r"\bCondos\b", s, re.I) and not re.search(r"\bInc\b", s, re.I):
            sc += 2
        if re.search(r"\bProperty Management\b", s, re.I):
            sc += 3
        if raw.isupper() or raw.islower():
            sc -= 3
        if "," in raw:
            sc -= 2
        if re.search(r"\s{2,}", raw):
            sc -= 2
        # Prefer "Breeze Bay" over "BreezeBay", and names with a legal suffix.
        if re.search(r"\bLLC\b", s, re.I):
            sc += 2
        compact = re.sub(r"[^a-z]+", "", sl)
        # Reward word breaks: "breeze bay club" beats "breezebay club".
        sc += s.count(" ") * 0.4
        sc += min(len(s), 40) / 20
        return (sc, len(s), s.count(" "))

    best = _flip_last_first(max(variants, key=score))
    best = re.sub(r"\s+", " ", best)
    best = re.sub(r"\bH\.\s*O\.\s*A\.?\b", "HOA", best, flags=re.I)
    best = re.sub(r"\bHOA\s+Inc\.?\b", "HOA", best, flags=re.I)
    best = proper_case_name(best)
    best = re.sub(r"\bHoa\b", "HOA", best)
    best = re.sub(r"\bLlc\b", "LLC", best)
    return best


def cluster_customers(rows: list[dict], log: QualityLog) -> tuple[list[dict], dict[int, str]]:
    """rows: raw mention dicts with keys name, email, phone, source, raw_label.

    Returns (customer_table, mention_index -> customer_id).
    """
    n = len(rows)
    uf = UnionFind(n)
    by_email: dict[str, int] = {}
    by_phone: dict[str, int] = {}
    by_name: dict[str, int] = {}
    by_core: dict[str, int] = {}

    for i, row in enumerate(rows):
        email, phone = row.get("email"), row.get("phone")
        name, core = row.get("norm_name") or "", row.get("core_name") or ""
        if email:
            if email in by_email:
                uf.union(i, by_email[email])
            else:
                by_email[email] = i
        if phone:
            if phone in by_phone:
                uf.union(i, by_phone[phone])
            else:
                by_phone[phone] = i
        if name:
            if name in by_name:
                uf.union(i, by_name[name])
            else:
                by_name[name] = i
        if core and core != name:
            if core in by_core:
                uf.union(i, by_core[core])
            else:
                by_core[core] = i

    # Second pass: fuzzy name match when last-name-ish tokens are close.
    # Only merge when there is a shared email/phone already in the same cluster
    # OR names are very close (typos / initials). We already merged exact keys;
    # this catches Gonzalez vs Gonzales with no shared email on that particular row.
    representatives = {}
    for i, row in enumerate(rows):
        name = row.get("norm_name") or row.get("core_name") or ""
        if not name:
            continue
        # Compare against earlier rows that are not yet in the same cluster.
        for j in range(i):
            other = rows[j].get("norm_name") or rows[j].get("core_name") or ""
            if not other:
                continue
            if uf.find(i) == uf.find(j):
                continue
            if names_similar(name, other) or names_similar(row.get("core_name") or "", rows[j].get("core_name") or ""):
                uf.union(i, j)

    groups: dict[int, list[int]] = defaultdict(list)
    for i in range(n):
        groups[uf.find(i)].append(i)

    customers: list[dict] = []
    mention_to_id: dict[int, str] = {}
    # Stable ids by first-seen order of each group's earliest mention.
    ordered_roots = sorted(groups, key=lambda r: min(groups[r]))
    for idx, root in enumerate(ordered_roots, start=1):
        members = groups[root]
        mentions = [rows[i] for i in members]
        cid = f"C{idx:03d}"
        # Pick a display name: prefer the longest original that is not all-caps,
        # and prefer forms that still include an entity suffix when present.
        emails = sorted({
            str(m["email"])
            for m in mentions
            if m.get("email") and not (isinstance(m.get("email"), float) and pd.isna(m.get("email")))
        })
        phones = sorted({
            str(m["phone"])
            for m in mentions
            if m.get("phone") and not (isinstance(m.get("phone"), float) and pd.isna(m.get("phone")))
        })
        variants = sorted({(m.get("raw_name") or "").strip() for m in mentions if (m.get("raw_name") or "").strip()})
        display = choose_display_name(mentions, emails)
        customers.append({
            "customer_id": cid,
            "customer_name": display,
            "email": emails[0] if emails else "",
            "phone": phones[0] if phones else "",
            "name_variants": " | ".join(variants),
            "variant_count": len(variants),
            "mention_count": len(mentions),
        })
        for i in members:
            mention_to_id[i] = cid
        if len(variants) > 1:
            log.add(
                "customer_merged",
                f"Merged {len(variants)} name variants into '{display}'",
                cid,
                variants=variants,
            )
    return customers, mention_to_id


# ---------------------------------------------------------------------------
# File loaders (skip preamble / locate headers)
# ---------------------------------------------------------------------------

def read_csv_with_preamble(path: Path) -> pd.DataFrame:
    """Skip leading non-header lines until a comma-separated header is found."""
    text = path.read_text(encoding="utf-8")
    lines = text.splitlines()
    header_idx = 0
    for i, line in enumerate(lines):
        if line.count(",") >= 3 and not line.lower().startswith("quickbooks") and "export" not in line.lower()[:40]:
            # A real header usually has no preamble phrasing and several columns.
            header_idx = i
            break
    from io import StringIO
    return pd.read_csv(StringIO("\n".join(lines[header_idx:])), dtype=str, keep_default_na=False)


def read_excel_sheet(path: Path, sheet: str) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name=sheet, header=None, dtype=object)
    header_idx = None
    for i, row in raw.iterrows():
        values = [str(v).strip() for v in row.tolist() if not is_blank(v)]
        joined = " ".join(values).lower()
        if "date" in joined and len(values) >= 4:
            header_idx = i
            break
    if header_idx is None:
        raise ValueError(f"Could not find a header row in {path.name} / {sheet}")
    header = [str(v).strip() if not is_blank(v) else f"col_{j}" for j, v in enumerate(raw.iloc[header_idx].tolist())]
    data = raw.iloc[header_idx + 1 :].copy()
    data.columns = header
    data = data.dropna(how="all")
    return data.reset_index(drop=True)


# ---------------------------------------------------------------------------
# Pipeline
# ---------------------------------------------------------------------------

def load_invoice_mentions(df: pd.DataFrame, log: QualityLog) -> pd.DataFrame:
    records = []
    for i, row in df.iterrows():
        rec_id = f"invoices_payments.csv:row{i+1}"
        inv_n = normalize_invoice_number(row.get("Invoice #"))
        inv_date = parse_date(row.get("Inv Date"), log, rec_id)
        due = parse_date(row.get("Due Date"), log, rec_id)
        if inv_date and inv_date.year == AS_OF.year + 1:
            fixed = inv_date.replace(year=AS_OF.year)
            if date(AS_OF.year, 1, 1) <= fixed <= AS_OF:
                log.add("date_fixed", f"Invoice date {inv_date.isoformat()} looked like a typo and was set to {fixed.isoformat()}", rec_id, invoice=inv_n)
                inv_date = fixed
        amount = parse_money(row.get("Amount"))
        tax = parse_money(row.get("Tax"))
        pay_date = parse_date(row.get("Payment Date"), log, rec_id)
        pay_amt = parse_money(row.get("Payment Amt"))
        service = canonicalize_service(row.get("Service"))
        email = normalize_email(row.get("Email"))
        phone = normalize_phone(row.get("Phone"))
        raw_name = str(row.get("Customer Name") or "").strip()
        if not inv_n:
            log.add("flagged", "Missing invoice number", rec_id)
        if is_blank(row.get("Inv Date")) or str(row.get("Inv Date")).strip() in BLANK_TOKENS or (
            not inv_date and not is_blank(row.get("Inv Date"))
        ):
            log.add("unparseable_date", f"Invoice date could not be parsed: {row.get('Inv Date')!r}", rec_id, invoice=inv_n)
        amount_raw = row.get("Amount")
        if (not is_blank(amount_raw) or (isinstance(amount_raw, str) and amount_raw.strip().lower() in GARBAGE_TOKENS)) and amount is None:
            log.add("unparseable_amount", f"Invoice amount could not be parsed: {amount_raw!r}", rec_id, invoice=inv_n)
        if row.get("Service") and not service:
            log.add("unknown_service", f"Unmapped service line: {row.get('Service')!r}", rec_id, invoice=inv_n)
        if inv_date and inv_date.year > AS_OF.year:
            log.add("flagged", f"Invoice date {inv_date.isoformat()} is after the sample year; typo suspected", rec_id, invoice=inv_n)
        if due and inv_date and due < inv_date:
            log.add("flagged", f"Due date {due.isoformat()} is before invoice date {inv_date.isoformat()}", rec_id, invoice=inv_n)
        records.append({
            "source_row": rec_id,
            "invoice_number": inv_n,
            "invoice_date": inv_date,
            "due_date": due,
            "raw_name": raw_name,
            "norm_name": normalize_name(raw_name),
            "core_name": core_name(raw_name),
            "email": email,
            "phone": phone,
            "service_raw": str(row.get("Service") or "").strip(),
            "service_line": service,
            "amount": amount,
            "tax": tax,
            "status_raw": str(row.get("Status") or "").strip(),
            "payment_date": pay_date,
            "payment_amount": pay_amt,
            "payment_method": str(row.get("Method") or "").strip(),
            "notes": str(row.get("Notes") or "").strip(),
            "raw_tuple": (
                str(row.get("Invoice #") or ""),
                str(row.get("Inv Date") or ""),
                str(row.get("Due Date") or ""),
                raw_name,
                str(row.get("Email") or ""),
                str(row.get("Phone") or ""),
                str(row.get("Service") or ""),
                str(row.get("Amount") or ""),
                str(row.get("Tax") or ""),
                str(row.get("Status") or ""),
                str(row.get("Payment Date") or ""),
                str(row.get("Payment Amt") or ""),
                str(row.get("Method") or ""),
                str(row.get("Notes") or ""),
            ),
        })
    return pd.DataFrame.from_records(records)


def dedup_invoice_rows(df: pd.DataFrame, log: QualityLog) -> pd.DataFrame:
    """Drop exact raw-row duplicates; keep the first occurrence."""
    if df.empty:
        return df
    seen = set()
    keep = []
    for idx, row in df.iterrows():
        key = row["raw_tuple"]
        if key in seen:
            log.add("duplicate_row_dropped", "Exact duplicate invoice-export row dropped", row["source_row"], invoice=row["invoice_number"])
            continue
        seen.add(key)
        keep.append(idx)
    return df.loc[keep].reset_index(drop=True)


def build_invoices_and_payments(df: pd.DataFrame, mention_offset: int, mention_to_id: dict, log: QualityLog) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Collapse payment-split rows into one invoice + many payments."""
    # mention indices for invoice df rows are mention_offset + original positional? We'll pass customer_id already.
    invoices = []
    payments = []
    grouped = df.groupby("invoice_number", dropna=False)
    used_ids = set()
    for inv_n, grp in grouped:
        if not inv_n:
            # Keep each unnumbered row as its own flagged invoice if it has any signal.
            for _, row in grp.iterrows():
                log.add("flagged", "Invoice skipped in clean invoice table (no invoice number)", row["source_row"])
            continue
        # Prefer a row with a parsed amount and date as the invoice master.
        scored = grp.copy()
        scored["_score"] = (
            scored["amount"].notna().astype(int) * 4
            + scored["invoice_date"].notna().astype(int) * 3
            + scored["service_line"].notna().astype(int)
            + scored["due_date"].notna().astype(int)
        )
        master = scored.sort_values("_score", ascending=False).iloc[0]
        # Conflicting amounts or dates across rows with the same invoice number.
        amounts = [a for a in grp["amount"].tolist() if a is not None and not (isinstance(a, float) and math.isnan(a))]
        dates = [d for d in grp["invoice_date"].tolist() if d is not None]
        if len(set(round(a, 2) for a in amounts)) > 1:
            log.add("flagged", f"Invoice {inv_n} has conflicting amounts {sorted(set(round(a, 2) for a in amounts))}", str(inv_n))
        if len(set(dates)) > 1:
            log.add("flagged", f"Invoice {inv_n} has conflicting invoice dates {sorted(set(d.isoformat() for d in dates))}", str(inv_n))
        invoices.append({
            "invoice_id": f"INV-{inv_n}",
            "invoice_number": inv_n,
            "invoice_date": master["invoice_date"],
            "due_date": master["due_date"] if pd.notna(master["due_date"]) else None,
            "customer_id": master["customer_id"],
            "customer_name": master["customer_name"],
            "service_line": master["service_line"] if pd.notna(master["service_line"]) else None,
            "amount": master["amount"] if pd.notna(master["amount"]) else None,
            "tax": master["tax"] if pd.notna(master["tax"]) else None,
            "notes": master["notes"],
            "source_rows": ";".join(grp["source_row"].tolist()),
        })
        used_ids.add(inv_n)
        pay_i = 1
        for _, row in grp.iterrows():
            if row["payment_amount"] is None or (isinstance(row["payment_amount"], float) and math.isnan(row["payment_amount"])):
                if is_blank(row.get("payment_method")) and row["payment_date"] is None:
                    continue
                # Method or date without amount: flag.
                if row["payment_date"] is not None or not is_blank(row.get("payment_method")):
                    log.add("flagged", f"Payment row on invoice {inv_n} is missing an amount", row["source_row"])
                continue
            method = re.sub(r"\s+", " ", str(row.get("payment_method") or "").strip())
            method_l = method.lower()
            if "credit" in method_l:
                method_c = "Credit memo"
            elif "ach" in method_l or "bank" in method_l:
                method_c = "ACH"
            elif "visa" in method_l or method_l == "card":
                method_c = "Visa"
            elif method_l in {"mc", "mastercard"}:
                method_c = "Mastercard"
            elif "cash" in method_l:
                method_c = "Cash"
            elif "chk" in method_l or "check" in method_l:
                method_c = "Check"
            elif not method:
                method_c = ""
            else:
                method_c = method
            payments.append({
                "payment_id": f"P-{inv_n}-{pay_i:02d}",
                "invoice_id": f"INV-{inv_n}",
                "invoice_number": inv_n,
                "customer_id": row["customer_id"],
                "payment_date": row["payment_date"],
                "amount": round(float(row["payment_amount"]), 2),
                "method": method_c,
                "source_row": row["source_row"],
            })
            pay_i += 1

    inv_df = pd.DataFrame(invoices)
    pay_df = pd.DataFrame(payments)
    if not inv_df.empty and not pay_df.empty:
        paid = pay_df.groupby("invoice_id")["amount"].sum().rename("amount_paid")
        inv_df = inv_df.merge(paid, on="invoice_id", how="left")
    else:
        inv_df["amount_paid"] = 0.0
    inv_df["amount_paid"] = inv_df["amount_paid"].fillna(0.0)
    balances = []
    statuses = []
    for _, row in inv_df.iterrows():
        amt = row["amount"]
        paid = float(row["amount_paid"])
        if amt is None or (isinstance(amt, float) and math.isnan(amt)):
            balances.append(None)
            statuses.append("Flagged")
            log.add("flagged", f"{row['invoice_id']} has no parseable amount; excluded from revenue and AR totals", row["invoice_id"])
            continue
        bal = round(float(amt) - paid, 2)
        balances.append(bal)
        if abs(bal) <= 0.01:
            statuses.append("Paid")
        elif paid == 0:
            statuses.append("Open")
        elif (amt > 0 and paid < amt) or (amt < 0 and paid > amt):
            statuses.append("Partial")
        else:
            statuses.append("Overpaid")
            log.add("flagged", f"{row['invoice_id']} is overpaid by ${paid - float(amt):.2f}", row["invoice_id"], overpay=round(paid - float(amt), 2))
        if row["invoice_date"] is None:
            log.add("flagged", f"{row['invoice_id']} has no parseable invoice date; excluded from monthly charts", row["invoice_id"])
    inv_df["balance"] = balances
    inv_df["status"] = statuses
    return inv_df, pay_df


def age_bucket(due: date | None, inv_date: date | None, balance: float | None, log: QualityLog, invoice_id: str) -> tuple[int | None, str]:
    if balance is None or abs(balance) <= 0.01:
        return 0, "Paid"
    ref = due or inv_date
    if ref is None:
        log.add("flagged", f"{invoice_id} cannot be aged (no due or invoice date)", invoice_id)
        return None, "Unknown"
    days = (AS_OF - ref).days
    if days <= 0:
        return days, "Current"
    if days <= 30:
        return days, "1–30"
    if days <= 60:
        return days, "31–60"
    if days <= 90:
        return days, "61–90"
    return days, "90+"


def chase_reason(row) -> str:
    days = row.get("days_past_due")
    name = row.get("customer_name") or "This customer"
    amount = row.get("balance") or 0
    notes = (row.get("notes") or "").lower()
    bucket = row.get("aging_bucket")
    if bucket == "90+":
        base = "Past 90 days — call and send a written balance confirmation."
    elif bucket == "61–90":
        base = "61–90 days past due — follow up this week."
    elif bucket == "31–60":
        base = "31–60 days past due — send a reminder and a copy of the invoice."
    elif bucket == "1–30":
        base = "Recently past due — a short reminder is enough to start."
    else:
        base = "Not yet due — keep on the radar."
    if "insurance" in notes:
        return f"{base} Notes mention waiting on insurance."
    if "no response" in notes or "voicemail" in notes or "follow-up" in notes:
        return f"{base} Earlier follow-up had no response."
    if amount >= 2000:
        return f"{base} This is one of the larger open balances."
    return base


def process_jobs(df: pd.DataFrame, customers: list[dict], invoices: pd.DataFrame, log: QualityLog) -> pd.DataFrame:
    # Build a name/email/phone lookup into existing customers, plus fuzzy.
    name_index = []
    for c in customers:
        name_index.append((c["customer_id"], c["customer_name"], normalize_name(c["customer_name"]), core_name(c["customer_name"])))
        for variant in (c.get("name_variants") or "").split(" | "):
            if variant:
                name_index.append((c["customer_id"], variant, normalize_name(variant), core_name(variant)))
    inv_by_n = {}
    if not invoices.empty:
        for _, row in invoices.iterrows():
            inv_by_n[row["invoice_number"]] = row

    def match_customer(raw_name: str) -> str | None:
        norm = normalize_name(raw_name)
        core = core_name(raw_name)
        for cid, _raw, n, c in name_index:
            if norm and norm == n:
                return cid
            if core and core == c:
                return cid
        for cid, _raw, n, c in name_index:
            if names_similar(norm, n) or (core and names_similar(core, c)):
                return cid
        return None

    rows = []
    for i, row in df.iterrows():
        rec = f"jobs_expenses.xlsx:Jobs:{i+1}"
        job_n = normalize_invoice_number(row.get("Job Number"))  # digits only, fine
        job_date = parse_date(row.get("Date"), log, rec)
        raw_name = str(row.get("Customer") or "").strip()
        service = canonicalize_service(row.get("Type"))
        inv_n = normalize_invoice_number(row.get("Invoice Ref"))
        labor = parse_money(row.get("Labor"))
        materials = parse_money(row.get("Materials"))
        hours = parse_money(row.get("Hours"))
        status_raw = str(row.get("Job Status") or "").strip()
        status_l = status_raw.lower()
        if "cancel" in status_l:
            status = "Cancelled"
        elif "complete" in status_l:
            status = "Completed"
        elif "schedul" in status_l:
            status = "Scheduled"
        elif not status_raw:
            status = "Unknown"
            log.add("flagged", f"Job {job_n or rec} has a blank status", rec)
        else:
            status = status_raw
        cid = match_customer(raw_name) if raw_name else None
        if raw_name and not cid:
            log.add("orphan_job", f"Job {job_n or rec} has no matching customer ('{raw_name}')", rec, name=raw_name)
        if not job_date and not is_blank(row.get("Date")):
            log.add("unparseable_date", f"Job date could not be parsed: {row.get('Date')!r}", rec)
        invoices_id = f"INV-{inv_n}" if inv_n else ""
        rows.append({
            "job_id": f"JOB-{job_n}" if job_n else f"JOB-ROW{i+1}",
            "job_date": job_date,
            "customer_id": cid or "",
            "customer_name": next((c["customer_name"] for c in customers if c["customer_id"] == cid), raw_name),
            "service_line": service,
            "invoice_id": invoices_id,
            "labor_cost": labor if labor is not None else 0.0,
            "materials_cost": materials if materials is not None else 0.0,
            "hours": hours,
            "tech": str(row.get("Tech") or "").strip(),
            "status": status,
            "notes": str(row.get("Notes") or "").strip(),
            "source_row": rec,
        })
    jobs = pd.DataFrame(rows)
    if not jobs.empty:
        jobs["direct_cost"] = (jobs["labor_cost"].fillna(0) + jobs["materials_cost"].fillna(0)).round(2)
    return jobs


def process_expenses(df: pd.DataFrame, log: QualityLog) -> pd.DataFrame:
    records = []
    seen = set()
    for i, row in df.iterrows():
        rec = f"jobs_expenses.xlsx:Expenses:{i+1}"
        dt = parse_date(row.get("Date"), log, rec)
        amount = parse_money(row.get("Amount"))
        vendor = re.sub(r"\s+", " ", str(row.get("Vendor") or "").strip())
        category = re.sub(r"\s+", " ", str(row.get("Category") or "").strip()).title()
        method = str(row.get("Payment Method") or "").strip()
        memo = str(row.get("Memo") or "").strip()
        ref = str(row.get("Ref") or "").strip()
        key = (dt, vendor.lower().strip(), round(amount, 2) if amount is not None else None, ref)
        if key in seen:
            log.add("duplicate_row_dropped", f"Duplicate expense row dropped ({vendor}, {dt}, {amount})", rec)
            continue
        seen.add(key)
        if amount is None:
            log.add("unparseable_amount", f"Expense amount could not be parsed: {row.get('Amount')!r}", rec)
        if dt is None and not is_blank(row.get("Date")):
            log.add("unparseable_date", f"Expense date could not be parsed: {row.get('Date')!r}", rec)
        records.append({
            "expense_id": ref if ref else f"E-ROW{i+1}",
            "expense_date": dt,
            "vendor": vendor.title() if vendor.isupper() else vendor,
            "category": category,
            "amount": amount,
            "method": method.title() if method else "",
            "memo": memo,
        })
    return pd.DataFrame.from_records(records)


def process_leads(df: pd.DataFrame, log: QualityLog) -> pd.DataFrame:
    records = []
    for i, row in df.iterrows():
        rec = f"crm_leads.csv:row{i+1}"
        lead_n = normalize_invoice_number(row.get("Lead ID"))
        created = parse_date(row.get("Created On"), log, rec)
        raw_name = str(row.get("Full Name") or "").strip()
        if "," in raw_name:
            parts = [p.strip() for p in raw_name.split(",", 1)]
            if len(parts) == 2:
                raw_name = f"{parts[1]} {parts[0]}"
        if raw_name.isupper() or raw_name.islower():
            raw_name = proper_case_name(raw_name)
        src_raw = str(row.get("Src") or "").strip()
        src = LEAD_SOURCE_MAP.get(src_raw.lower().strip(), proper_case_name(src_raw) if src_raw else "")
        if not src and src_raw:
            log.add("flagged", f"Lead source not standardized: {src_raw!r}", rec)
        status_raw = str(row.get("Lead Status") or "").strip()
        status = LEAD_STATUS_MAP.get(status_raw.lower().strip(), "")
        if status_raw and not status:
            log.add("flagged", f"Lead status not standardized: {status_raw!r}", rec)
            status = status_raw
        if not status:
            status = "Unknown"
            log.add("flagged", "Lead is missing a status", rec)
        est = parse_money(row.get("Est. Value"))
        if not is_blank(row.get("Est. Value")) and est is None:
            log.add("unparseable_amount", f"Lead estimated value could not be parsed: {row.get('Est. Value')!r}", rec)
        if created is None and not is_blank(row.get("Created On")):
            log.add("unparseable_date", f"Lead created date could not be parsed: {row.get('Created On')!r}", rec)
        records.append({
            "lead_id": f"L-{lead_n}" if lead_n else f"L-ROW{i+1}",
            "created_date": created,
            "contact_name": raw_name,
            "source": src or "",
            "status": status,
            "estimated_value": est,
            "notes": str(row.get("Notes") or "").strip(),
        })
    return pd.DataFrame.from_records(records)


def compute_kpis(invoices: pd.DataFrame, payments: pd.DataFrame, jobs: pd.DataFrame, expenses: pd.DataFrame, customers: list[dict], log: QualityLog) -> dict:
    inv = invoices.copy()
    inv["invoice_date"] = pd.to_datetime(inv["invoice_date"], errors="coerce")
    usable = inv[inv["amount"].notna() & inv["invoice_date"].notna()].copy()
    usable["amount"] = usable["amount"].astype(float)
    usable["month"] = usable["invoice_date"].dt.to_period("M").astype(str)

    jobs_u = jobs.copy()
    jobs_u["job_date"] = pd.to_datetime(jobs_u["job_date"], errors="coerce")
    completed = jobs_u[(jobs_u["status"] == "Completed") & jobs_u["job_date"].notna()].copy()
    completed["month"] = completed["job_date"].dt.to_period("M").astype(str)

    # Revenue by month / service from invoices; direct cost from completed jobs.
    rev_month = usable.groupby("month")["amount"].sum().sort_index()
    cost_month = completed.groupby("month")["direct_cost"].sum().sort_index()
    months = sorted(set(rev_month.index) | set(cost_month.index))
    monthly = []
    for m in months:
        rev = float(rev_month.get(m, 0) or 0)
        cost = float(cost_month.get(m, 0) or 0)
        gp = rev - cost
        monthly.append({
            "month": m,
            "label": datetime.strptime(m + "-01", "%Y-%m-%d").strftime("%b %Y"),
            "revenue": round(rev, 2),
            "direct_cost": round(cost, 2),
            "gross_profit": round(gp, 2),
            "margin_pct": round(gp / rev, 4) if rev else None,
        })

    rev_line = usable.groupby("service_line")["amount"].sum()
    cost_line = completed.groupby("service_line")["direct_cost"].sum()
    lines = []
    for line in CANONICAL_SERVICES:
        rev = float(rev_line.get(line, 0) or 0)
        cost = float(cost_line.get(line, 0) or 0)
        gp = rev - cost
        lines.append({
            "service_line": line,
            "revenue": round(rev, 2),
            "direct_cost": round(cost, 2),
            "gross_profit": round(gp, 2),
            "margin_pct": round(gp / rev, 4) if rev else None,
            "invoice_count": int((usable["service_line"] == line).sum()),
        })

    # AR
    ar = inv[(inv["balance"].notna()) & (inv["balance"].astype(float).abs() > 0.01) & (inv["status"].isin(["Open", "Partial"]))].copy()
    days_list, buckets = [], []
    for _, row in ar.iterrows():
        due = row["due_date"] if pd.notna(row["due_date"]) else None
        if isinstance(due, pd.Timestamp):
            due = due.date()
        inv_d = row["invoice_date"]
        if isinstance(inv_d, pd.Timestamp):
            inv_d = inv_d.date()
        elif pd.isna(inv_d):
            inv_d = None
        d, b = age_bucket(due, inv_d, float(row["balance"]), log, row["invoice_id"])
        days_list.append(d)
        buckets.append(b)
    ar["days_past_due"] = days_list
    ar["aging_bucket"] = buckets
    ar["chase_reason"] = ar.apply(chase_reason, axis=1)

    aging_order = ["Current", "1–30", "31–60", "61–90", "90+"]
    aging = []
    for b in aging_order:
        subset = ar[ar["aging_bucket"] == b]
        aging.append({
            "bucket": b,
            "balance": round(float(subset["balance"].sum()), 2) if len(subset) else 0.0,
            "invoice_count": int(len(subset)),
        })

    # Chase first: roll past-due invoices up by customer (who to call), ranked by balance * oldest days.
    past = ar[ar["aging_bucket"].isin(["1–30", "31–60", "61–90", "90+"])].copy()
    chase = []
    if not past.empty:
        rows_out = []
        for cid, grp in past.groupby("customer_id", dropna=False):
            grp = grp.sort_values(["days_past_due", "balance"], ascending=False)
            top = grp.iloc[0]
            total = round(float(grp["balance"].sum()), 2)
            days = int(top["days_past_due"]) if pd.notna(top["days_past_due"]) else 0
            open_all = ar[ar["customer_id"] == cid] if "customer_id" in ar.columns else grp
            total_open = round(float(open_all["balance"].sum()), 2)
            invoices_ids = ", ".join(grp["invoice_id"].tolist())
            reason = chase_reason(top)
            if len(grp) > 1:
                reason = f"{len(grp)} past-due invoices, ${total:,.2f} total. " + reason
            inv_d = top["invoice_date"]
            if isinstance(inv_d, pd.Timestamp):
                inv_d = inv_d.date()
            due = top["due_date"]
            if isinstance(due, pd.Timestamp):
                due = due.date()
            rows_out.append({
                "customer_id": cid,
                "customer_name": top["customer_name"],
                "invoice_id": top["invoice_id"],
                "invoice_ids": invoices_ids,
                "service_line": top["service_line"] if pd.notna(top["service_line"]) else "",
                "invoice_date": inv_d.isoformat() if inv_d else "",
                "due_date": due.isoformat() if due else "",
                "balance": total,
                "total_open_ar": total_open,
                "days_past_due": days if days else None,
                "aging_bucket": top["aging_bucket"],
                "past_due_count": int(len(grp)),
                "reason": reason,
                "score": total * max(days, 1),
            })
        rows_out.sort(key=lambda r: (r["score"], r["balance"]), reverse=True)
        chase = [{k: v for k, v in r.items() if k != "score"} for r in rows_out[:8]]

    # Cash
    pay = payments.copy()
    pay["payment_date"] = pd.to_datetime(pay["payment_date"], errors="coerce")
    cash_in_m = pay[pay["payment_date"].notna()].groupby(pay["payment_date"].dt.to_period("M").astype(str))["amount"].sum()
    exp = expenses.copy()
    exp["expense_date"] = pd.to_datetime(exp["expense_date"], errors="coerce")
    overhead_m = exp[exp["expense_date"].notna() & exp["amount"].notna()].groupby(exp["expense_date"].dt.to_period("M").astype(str))["amount"].sum()
    # Job direct costs treated as paid in the job month (documented simplification).
    cash_months = sorted(set(months) | set(cash_in_m.index) | set(overhead_m.index))
    cash = []
    for m in cash_months:
        cin = float(cash_in_m.get(m, 0) or 0)
        cjob = float(cost_month.get(m, 0) or 0)
        covh = float(overhead_m.get(m, 0) or 0)
        cout = cjob + covh
        cash.append({
            "month": m,
            "label": datetime.strptime(m + "-01", "%Y-%m-%d").strftime("%b %Y") if re.fullmatch(r"\d{4}-\d{2}", m) else m,
            "cash_in": round(cin, 2),
            "job_costs": round(cjob, 2),
            "overhead": round(covh, 2),
            "cash_out": round(cout, 2),
            "net_cash": round(cin - cout, 2),
        })

    total_rev = round(float(usable["amount"].sum()), 2)
    total_cost = round(float(completed["direct_cost"].sum()), 2)
    total_gp = round(total_rev - total_cost, 2)
    open_ar = round(float(ar["balance"].sum()) if len(ar) else 0.0, 2)
    cash_in_total = round(float(pay["amount"].sum()) if len(pay) else 0.0, 2)
    cash_out_total = round(float(completed["direct_cost"].sum()) + float(exp["amount"].fillna(0).sum()), 2)

    return {
        "as_of": AS_OF.isoformat(),
        "company": "Lantern Key Pool & Air",
        "disclaimer": "SAMPLE: synthetic data for a fictional business. Not a real company and not real results.",
        "totals": {
            "revenue": total_rev,
            "direct_cost": total_cost,
            "gross_profit": total_gp,
            "margin_pct": round(total_gp / total_rev, 4) if total_rev else None,
            "open_ar": open_ar,
            "cash_in": cash_in_total,
            "cash_out": cash_out_total,
            "net_cash": round(cash_in_total - cash_out_total, 2),
            "invoice_count": int(len(usable)),
            "customer_count": len(customers),
        },
        "monthly": monthly,
        "service_lines": lines,
        "aging": aging,
        "chase_first": chase,
        "cash_monthly": cash,
        "ar_invoices": [
            {
                "invoice_id": r["invoice_id"],
                "customer_name": r["customer_name"],
                "service_line": r["service_line"] if pd.notna(r["service_line"]) else "",
                "invoice_date": (r["invoice_date"].date().isoformat() if isinstance(r["invoice_date"], pd.Timestamp) else (r["invoice_date"].isoformat() if r["invoice_date"] else "")),
                "due_date": (r["due_date"].date().isoformat() if isinstance(r["due_date"], pd.Timestamp) else (r["due_date"].isoformat() if pd.notna(r["due_date"]) and r["due_date"] else "")),
                "amount": round(float(r["amount"]), 2) if pd.notna(r["amount"]) else None,
                "amount_paid": round(float(r["amount_paid"]), 2),
                "balance": round(float(r["balance"]), 2),
                "days_past_due": int(r["days_past_due"]) if pd.notna(r["days_past_due"]) else None,
                "aging_bucket": r["aging_bucket"],
                "status": r["status"],
            }
            for _, r in ar.sort_values(["days_past_due", "balance"], ascending=False).iterrows()
        ],
    }


def write_quality_log(log: QualityLog, path_md: Path, path_json: Path, summary: dict) -> None:
    lines = [
        "# Data quality log",
        "",
        "SAMPLE: synthetic data for a fictional business (Lantern Key Pool & Air).",
        f"As-of date: {AS_OF.isoformat()}. Generated by `clean_data.py`.",
        "",
        "## Summary",
        "",
    ]
    for key, label in [
        ("raw_invoice_rows", "Raw invoice-export rows read"),
        ("duplicate_invoice_rows_dropped", "Exact duplicate invoice-export rows dropped"),
        ("clean_invoices", "Clean invoices written"),
        ("clean_payments", "Clean payments written"),
        ("customer_variants_in", "Customer name strings seen in the invoice export"),
        ("customers_after_merge", "Customers after merge"),
        ("customer_groups_merged", "Customer groups that had more than one name variant"),
        ("dates_standardized", "Dates standardized (invoice, due, and payment dates parsed)"),
        ("amounts_parsed", "Currency amounts parsed (invoice and payment)"),
        ("service_lines_standardized", "Service lines standardized"),
        ("unparseable_dates", "Dates that could not be parsed"),
        ("unparseable_amounts", "Amounts that could not be parsed"),
        ("rows_flagged", "Rows or invoices flagged for review"),
        ("orphan_jobs", "Jobs with no matching customer"),
        ("duplicate_expense_rows_dropped", "Duplicate expense rows dropped"),
        ("clean_jobs", "Clean jobs written"),
        ("clean_expenses", "Clean expenses written"),
        ("clean_leads", "Clean leads written"),
    ]:
        if key in summary:
            lines.append(f"- **{label}:** {summary[key]}")
    lines += ["", "## Event counts", ""]
    for kind in sorted(log.counts):
        lines.append(f"- `{kind}`: {log.counts[kind]}")
    lines += ["", "## Events", ""]
    for ev in log.events:
        rec = f" ({ev.record})" if ev.record else ""
        lines.append(f"- **{ev.kind}**{rec}: {ev.message}")
    path_md.write_text("\n".join(lines) + "\n", encoding="utf-8")
    path_json.write_text(
        json.dumps({"summary": summary, "counts": dict(log.counts), "events": [asdict(e) for e in log.events]}, indent=2, default=str),
        encoding="utf-8",
    )


def _iso(v):
    if v is None or (isinstance(v, float) and math.isnan(v)) or pd.isna(v):
        return ""
    if isinstance(v, pd.Timestamp):
        return v.date().isoformat()
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    return v


def main() -> None:
    CLEAN.mkdir(parents=True, exist_ok=True)
    log = QualityLog()

    inv_raw = read_csv_with_preamble(RAW / "invoices_payments.csv")
    jobs_raw = read_excel_sheet(RAW / "jobs_expenses.xlsx", "Jobs Export")
    exp_raw = read_excel_sheet(RAW / "jobs_expenses.xlsx", "Expenses")
    leads_raw = read_csv_with_preamble(RAW / "crm_leads.csv")

    raw_invoice_rows = len(inv_raw)
    mentions = load_invoice_mentions(inv_raw, log)
    dates_std = int(mentions["invoice_date"].notna().sum() + mentions["due_date"].notna().sum() + mentions["payment_date"].notna().sum())
    amounts_parsed = int(mentions["amount"].notna().sum() + mentions["payment_amount"].notna().sum())
    services_std = int(mentions["service_line"].notna().sum())

    before = len(mentions)
    mentions = dedup_invoice_rows(mentions, log)
    dups_dropped = before - len(mentions)

    # Customer mentions from the de-duplicated invoice export (plus jobs/leads later via matching).
    mention_rows = []
    for _, row in mentions.iterrows():
        email = row["email"] if pd.notna(row["email"]) else None
        phone = row["phone"] if pd.notna(row["phone"]) else None
        mention_rows.append({
            "raw_name": row["raw_name"] if pd.notna(row["raw_name"]) else "",
            "norm_name": row["norm_name"] if pd.notna(row["norm_name"]) else "",
            "core_name": row["core_name"] if pd.notna(row["core_name"]) else "",
            "email": email or None,
            "phone": phone or None,
            "source": row["source_row"],
        })
    unique_name_strings = {m["raw_name"] for m in mention_rows if m["raw_name"]}
    customers, mention_to_id = cluster_customers(mention_rows, log)
    cust_by_id = {c["customer_id"]: c for c in customers}
    mentions = mentions.copy()
    mentions["customer_id"] = [mention_to_id[i] for i in range(len(mentions))]
    mentions["customer_name"] = mentions["customer_id"].map(lambda cid: cust_by_id[cid]["customer_name"])

    invoices, payments = build_invoices_and_payments(mentions, 0, mention_to_id, log)
    jobs = process_jobs(jobs_raw, customers, invoices, log)
    expenses = process_expenses(exp_raw, log)
    leads = process_leads(leads_raw, log)

    kpis = compute_kpis(invoices, payments, jobs, expenses, customers, log)

    # Write clean tables
    def dump(df: pd.DataFrame, name: str, date_cols: list[str]) -> None:
        out = df.copy()
        for c in date_cols:
            if c in out.columns:
                out[c] = out[c].map(_iso)
        out.to_csv(CLEAN / name, index=False)

    dump(pd.DataFrame(customers), "customers.csv", [])
    dump(invoices.drop(columns=["source_rows"], errors="ignore"), "invoices.csv", ["invoice_date", "due_date"])
    dump(payments, "payments.csv", ["payment_date"])
    dump(jobs, "jobs.csv", ["job_date"])
    dump(expenses, "expenses.csv", ["expense_date"])
    dump(leads, "leads.csv", ["created_date"])
    if kpis["ar_invoices"]:
        pd.DataFrame(kpis["ar_invoices"]).to_csv(CLEAN / "ar_open.csv", index=False)
    if kpis["chase_first"]:
        pd.DataFrame(kpis["chase_first"]).to_csv(CLEAN / "ar_chase_first.csv", index=False)
    pd.DataFrame(kpis["monthly"]).to_csv(CLEAN / "monthly_kpis.csv", index=False)
    pd.DataFrame(kpis["service_lines"]).to_csv(CLEAN / "service_line_kpis.csv", index=False)
    pd.DataFrame(kpis["cash_monthly"]).to_csv(CLEAN / "cash_monthly.csv", index=False)

    groups_merged = sum(1 for c in customers if c["variant_count"] > 1)
    flagged = log.count("flagged") + log.count("unparseable_date") + log.count("unparseable_amount") + log.count("orphan_job") + log.count("unknown_service")
    summary = {
        "raw_invoice_rows": raw_invoice_rows,
        "duplicate_invoice_rows_dropped": dups_dropped,
        "clean_invoices": int(len(invoices)),
        "clean_payments": int(len(payments)),
        "customer_variants_in": len(unique_name_strings),
        "customers_after_merge": len(customers),
        "customer_groups_merged": groups_merged,
        "dates_standardized": dates_std,
        "amounts_parsed": amounts_parsed,
        "service_lines_standardized": services_std,
        "unparseable_dates": log.count("unparseable_date"),
        "unparseable_amounts": log.count("unparseable_amount"),
        "rows_flagged": flagged,
        "orphan_jobs": log.count("orphan_job"),
        "duplicate_expense_rows_dropped": sum(1 for ev in log.events if ev.kind == "duplicate_row_dropped" and "expense" in ev.message.lower()),
        "clean_jobs": int(len(jobs)),
        "clean_expenses": int(len(expenses)),
        "clean_leads": int(len(leads)),
    }
    kpis["quality"] = summary
    kpis["quality_event_counts"] = dict(log.counts)

    (CLEAN / "kpis.json").write_text(json.dumps(kpis, indent=2, default=str), encoding="utf-8")
    write_quality_log(log, CLEAN / "data_quality_log.md", CLEAN / "data_quality_log.json", summary)

    print("Clean tables written to", CLEAN)
    for k, v in summary.items():
        print(f"  {k}: {v}")
    t = kpis["totals"]
    print("  --- KPIs ---")
    print(f"  revenue: ${t['revenue']:,.2f}")
    print(f"  gross_profit: ${t['gross_profit']:,.2f} ({(t['margin_pct'] or 0)*100:.1f}%)")
    print(f"  open_ar: ${t['open_ar']:,.2f}")
    print(f"  cash_in: ${t['cash_in']:,.2f}  cash_out: ${t['cash_out']:,.2f}")
    print("  aging:", kpis["aging"])
    print("  chase first:")
    for c in kpis["chase_first"]:
        print(f"    {c['customer_name']}: ${c['balance']:,.2f}  {c['days_past_due']}d  {c['aging_bucket']}")


if __name__ == "__main__":
    main()
