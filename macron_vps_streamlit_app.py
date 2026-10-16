"""Macron / VPS / Paymeter / KYC reconciliation Streamlit app.

Run:
    python -m pip install -r requirements.txt
    streamlit run macron_vps_streamlit_app.py

Uploaded customer files are processed in memory; no external API calls.
"""

from __future__ import annotations

import csv
import io
import math
import posixpath
import re
import zipfile
from collections import Counter, defaultdict
from datetime import date, datetime, time, timedelta, timezone
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any
from xml.etree import ElementTree as ET

import xlsxwriter


MAX_FILE_BYTES = 30 * 1024 * 1024
MAX_XLSX_UNCOMPRESSED = 160 * 1024 * 1024
MAX_ROWS = 200_000
FEE_CENTS = 10_000  # NGN 100.00
MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
DOCREL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

POSTING = [
    "transaction_date", "customer_name", "account_number", "cash_in",
    "cash_out", "interest", "source",
]
AUDIT = [
    "status", "confidence", "matching_method", "review_notes", "duplicate_token_candidate",
    "identity_status", "Macron_account_KYC", "virtual_account_KYC",
    "Macron_source_row", "VPS_source_row", "Paymeter_source_row",
    "Paymeter_reference", "Paymeter_input_amount", "Paymeter_token_amount",
    "PayBill_group", "VPS_gross_NGN", "VPS_actual_bank_charge_NGN",
    "VPS_reported_bank_charge_NGN", "VPS_charge_difference_NGN",
    "VPS_settled_NGN", "expected_token_NGN", "token_variance_NGN",
]
MACRON_FIELDS = [
    "DATE & TIME", "TRANSACTION ID", "METER NUMBER", "ACCOUNT NUMBER", "ACTIVITY",
    "AMOUNT", "STATUS", "REFERENCE ID", "SOURCE",
]
VPS_FIELDS = [
    "id", "session_id", "settlement_ref", "virtual_acct_no", "cba_tran_time",
    "created_at", "transaction_amount_minor", "settled_amount_minor",
    "charge_amount_minor", "source_acct_name", "narration",
]


def norm(value: Any) -> str:
    """Preserve identifier digits and leading zeros where represented as text."""
    if value is None:
        return ""
    s = str(value).strip()
    if re.fullmatch(r"\d+\.0+", s):
        return s.split(".", 1)[0]
    return s


def reference(value: Any) -> str:
    return norm(value).casefold()


def as_cents(value: Any) -> int | None:
    if value is None or str(value).strip() == "":
        return None
    try:
        d = Decimal(str(value).replace(",", ""))
        if not d.is_finite():
            return None
        return int((d * 100).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    except (InvalidOperation, ValueError):
        return None


def ngn(cents: int | None) -> float | None:
    return round(cents / 100, 2) if cents is not None else None


def first_datetime(value: Any, *, utc_to_lagos: bool = False) -> datetime | None:
    if isinstance(value, datetime):
        return value.replace(tzinfo=None)
    if isinstance(value, date):
        return datetime.combine(value, time.min)
    if isinstance(value, (int, float)) and 20000 < float(value) < 90000:
        return datetime(1899, 12, 30) + timedelta(days=float(value))
    s = norm(value)
    if re.fullmatch(r"[0-9]{5}(?:\.[0-9]+)?", s) and 20000 < float(s) < 90000:
        return datetime(1899, 12, 30) + timedelta(days=float(s))
    if not s:
        return None
    try:
        iso = datetime.fromisoformat(s.replace("Z", "+00:00"))
        if iso.tzinfo is not None:
            return iso.astimezone(timezone(timedelta(hours=1))).replace(tzinfo=None)
        if utc_to_lagos:
            return iso + timedelta(hours=1)
        return iso
    except ValueError:
        pass
    for fmt in (
        "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y %H:%M:%S.%f",
        "%d/%m/%Y", "%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d", "%d-%b-%y", "%d-%b-%Y",
    ):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def safe_xml(data: bytes) -> ET.Element:
    if b"<!DOCTYPE" in data.upper() or b"<!ENTITY" in data.upper():
        raise ValueError("Unsafe XML declarations in uploaded workbook")
    return ET.fromstring(data)


def read_xlsx(content: bytes) -> tuple[list[dict], list[str]]:
    """Read first worksheet XLSX without Excel macros, formulas, or external links."""
    if len(content) > MAX_FILE_BYTES:
        raise ValueError("Workbook exceeds 30 MB maximum")
    try:
        z = zipfile.ZipFile(io.BytesIO(content))
    except (zipfile.BadZipFile, OSError) as exc:
        raise ValueError("Invalid XLSX file") from exc
    with z:
        members = z.infolist()
        if sum(i.file_size for i in members) > MAX_XLSX_UNCOMPRESSED:
            raise ValueError("Workbook too large when decompressed")
        paths = set(z.namelist())
        for required in ("xl/workbook.xml", "xl/_rels/workbook.xml.rels"):
            if required not in paths:
                raise ValueError(f"Invalid Excel workbook: missing {required}")
        workbook = safe_xml(z.read("xl/workbook.xml"))
        relations = safe_xml(z.read("xl/_rels/workbook.xml.rels"))
        rels = {r.get("Id"): r.get("Target") for r in relations}
        sheets = workbook.findall(".//" + MAIN + "sheet")
        if not sheets:
            raise ValueError("No worksheets found")
        target = rels.get(sheets[0].get(DOCREL + "id"), "")
        if not target or "://" in target:
            raise ValueError("Worksheet target is invalid")
        target = target.lstrip("/") if target.startswith("/") else "xl/" + target
        target = posixpath.normpath(target)
        if not target.startswith("xl/worksheets/") or target not in paths:
            raise ValueError("Worksheet target is not a valid XLSX sheet")
        shared = []
        if "xl/sharedStrings.xml" in paths:
            root = safe_xml(z.read("xl/sharedStrings.xml"))
            shared = ["".join(t.text or "" for t in e.iter(MAIN + "t"))
                      for e in root.findall(MAIN + "si")]
        root = safe_xml(z.read(target))
        parsed = []
        for xrow in root.findall(".//" + MAIN + "sheetData/" + MAIN + "row"):
            cells: dict[int, str] = {}
            for cell in xrow.findall(MAIN + "c"):
                coord = cell.get("r", "")
                letters = re.match(r"^[A-Z]+", coord)
                if not letters:
                    continue
                col = 0
                for c in letters.group(0):
                    col = 26 * col + ord(c) - 64
                col -= 1
                val = cell.find(MAIN + "v")
                item = "" if val is None or val.text is None else val.text
                kind = cell.get("t", "")
                if kind == "s" and item:
                    index = int(item)
                    if index < 0 or index >= len(shared):
                        raise ValueError("Invalid shared string index")
                    item = shared[index]
                elif kind == "inlineStr":
                    item = "".join(t.text or "" for t in cell.iter(MAIN + "t"))
                cells[col] = item
            parsed.append((int(xrow.get("r", "0")), cells))
            if len(parsed) > MAX_ROWS + 1:
                raise ValueError("Too many rows in uploaded worksheet")
    if not parsed:
        raise ValueError("First worksheet is empty")
    head = {col: norm(val) for col, val in parsed[0][1].items() if norm(val)}
    if not head:
        raise ValueError("Column headers are missing")
    headings = list(head.values())
    if len(headings) != len(set(headings)):
        raise ValueError("Duplicate column headings in XLSX")
    rows = []
    for rowno, cells in parsed[1:]:
        if not any(norm(cells.get(i)) for i in head):
            continue
        d = {name: cells.get(i, "") for i, name in head.items()}
        d["_row"] = rowno
        rows.append(d)
    return rows, headings


def read_paymeter(content: bytes) -> tuple[list[dict], list[str]]:
    if len(content) > MAX_FILE_BYTES:
        raise ValueError("Paymeter CSV exceeds 30 MB")
    try:
        f = io.StringIO(content.decode("utf-8-sig"), newline="")
    except UnicodeDecodeError as exc:
        raise ValueError("Paymeter CSV must use UTF-8 encoding") from exc
    rd = csv.reader(f)
    try:
        full_heads = next(rd)
    except StopIteration as exc:
        raise ValueError("Empty Paymeter CSV") from exc
    usable = [i for i, col in enumerate(full_heads) if col.strip()]
    head = [full_heads[i].strip() for i in usable]
    if len(head) != len(set(head)):
        raise ValueError("Duplicate headers in Paymeter CSV")
    rows = []
    for source_row, fields in enumerate(rd, 2):
        if not any(f.strip() for f in fields):
            continue
        # Never silently shift Paymeter columns when a malformed CSV has extra commas.
        if len(fields) > len(full_heads):
            raise ValueError(f"Paymeter CSV row {source_row} contains too many columns; correct its quoting")
        fields += [""] * (len(full_heads) - len(fields))
        row = {full_heads[i].strip(): fields[i] for i in usable}
        # Check stable anchor fields: commas in unquoted addresses can otherwise
        # shift values without increasing row length (the CSV has blank tail columns).
        if ("Reference" in row and norm(row["Reference"])
                and not re.fullmatch(r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}", norm(row["Reference"]))):
            raise ValueError(f"Paymeter CSV row {source_row}: reference is not in the expected UUID column; check CSV quoting")
        if as_cents(row.get("Input Amount")) is None or as_cents(row.get("Transaction Amount")) is None:
            raise ValueError(f"Paymeter CSV row {source_row}: amount columns are missing or shifted")
        row["_row"] = source_row
        rows.append(row)
        if len(rows) > MAX_ROWS:
            raise ValueError("Paymeter CSV exceeds allowed row count")
    return rows, head


def require(rows: list[dict], headers: list[str], needed: list[str], label: str) -> None:
    missing = sorted(set(needed) - set(headers))
    if missing:
        raise ValueError(f"{label} missing required columns: {', '.join(missing)}")
    if not rows:
        raise ValueError(f"{label} has no transaction rows")


def prepare_inputs(macron_data: bytes, paymeter_data: bytes, vps_data: bytes, kyc_data: bytes) -> dict:
    m, mh = read_xlsx(macron_data)
    v, vh = read_xlsx(vps_data)
    k, kh = read_xlsx(kyc_data)
    p, ph = read_paymeter(paymeter_data)
    require(m, mh, ["REFERENCE ID", "ACCOUNT NUMBER", "ACTIVITY", "AMOUNT", "DATE & TIME"], "Macron")
    require(v, vh, ["virtual_acct_no", "transaction_amount_minor", "settled_amount_minor"], "VPS")
    require(k, kh, ["PAYMETER - Account Number", "PAYMETER - Customer Name", "MACRON ACCOUNT NUMBER"], "KYC")
    require(p, ph, ["Reference", "Account Number", "Input Amount", "Transaction Amount", "Created At"], "Paymeter")
    return {"m": m, "v": v, "k": k, "p": p, "mh": mh, "vh": vh, "kh": kh, "ph": ph}


def reconcile(data: dict, fee_ngn: int = 100, max_delay_hours: int = 36,
              high_confidence_minutes: int = 30) -> dict:
    """Apply exact reference bridge first; then conservative account/amount/time matches."""
    m, v, p, k = (data[name] for name in ("m", "v", "p", "k"))
    fee = fee_ngn * 100
    k_va: dict[str, dict] = {}
    k_mac: dict[str, dict] = {}
    conflict_virtual: set[str] = set()
    conflict_mac: set[str] = set()
    for row in k:
        va = norm(row.get("PAYMETER - Account Number"))
        mac = norm(row.get("MACRON ACCOUNT NUMBER"))
        if not va:
            continue
        entry = {"va": va, "mac": mac, "name": norm(row.get("PAYMETER - Customer Name")),
                 "meter": norm(row.get("METER NUMBER")), "row": row["_row"]}
        old = k_va.get(va)
        if old and old["mac"] != mac:
            conflict_virtual.add(va)
        else:
            k_va.setdefault(va, entry)
        if mac:
            old = k_mac.get(mac)
            if old and old["va"] != va:
                conflict_mac.add(mac)
            else:
                k_mac.setdefault(mac, entry)
    for va in conflict_virtual:
        bad = k_va.pop(va, None)
        if bad:
            k_mac.pop(bad["mac"], None)
    for mac in conflict_mac:
        bad = k_mac.pop(mac, None)
        if bad:
            k_va.pop(bad["va"], None)

    p_ref: dict[str, list[int]] = defaultdict(list)
    for j, row in enumerate(p):
        row["_va"] = norm(row.get("Account Number"))
        row["_gross"] = as_cents(row.get("Input Amount"))
        row["_token"] = as_cents(row.get("Transaction Amount"))
        row["_date"] = first_datetime(row.get("Created At"))
        row["_reference"] = reference(row.get("Reference"))
        if row["_reference"]:
            p_ref[row["_reference"]].append(j)

    for j, row in enumerate(m):
        row["_date"] = first_datetime(row.get("DATE & TIME"))
        row["_mac"] = norm(row.get("ACCOUNT NUMBER"))
        row["_amount"] = as_cents(row.get("AMOUNT"))
        row["_activity"] = norm(row.get("ACTIVITY")).upper()
        row["_reference"] = reference(row.get("REFERENCE ID"))
        row["_pidx"] = None
        if len(p_ref.get(row["_reference"], [])) == 1:
            row["_pidx"] = p_ref[row["_reference"]][0]
        owner = k_mac.get(row["_mac"])
        row["_va"] = owner["va"] if owner else ""
        row["_identity"] = "KYC" if owner else "UNMAPPED"
        if row["_pidx"] is not None:
            pv = p[row["_pidx"]]["_va"]
            if row["_va"] and pv and pv != row["_va"]:
                row["_identity"] = "CONFLICT"
            elif not row["_va"] and pv and pv not in conflict_virtual:
                # A reference link is a provisional account bridge, not full KYC.
                row["_va"] = pv
                row["_identity"] = "REFERENCE_DERIVED"

    for j, row in enumerate(v):
        row["_va"] = norm(row.get("virtual_acct_no"))
        row["_gross"] = as_cents(row.get("transaction_amount_minor"))
        row["_settled"] = as_cents(row.get("settled_amount_minor"))
        row["_charge"] = as_cents(row.get("charge_amount_minor"))
        row["_date"] = first_datetime(row.get("cba_tran_time"))
        if row["_date"] is None:
            row["_date"] = first_datetime(row.get("created_at"), utc_to_lagos=True)

    by_va_amount: dict[tuple[str, int], list[int]] = defaultdict(list)
    by_va: dict[str, list[int]] = defaultdict(list)
    for j, row in enumerate(v):
        by_va[row["_va"]].append(j)
        if row["_gross"] is not None and row["_va"]:
            by_va_amount[(row["_va"], row["_gross"])].append(j)

    matched_m: dict[int, int] = {}
    matched_v: dict[int, int] = {}
    match_meta: dict[int, dict] = {}
    max_minutes = max_delay_hours * 60

    def assign(mi: int, vi: int, method: str, confidence: str, notes: str, pidx: int | None = None) -> None:
        if mi in matched_m or vi in matched_v:
            raise AssertionError("Transaction assigned twice")
        matched_m[mi] = vi
        matched_v[vi] = mi
        match_meta[mi] = {"method": method, "confidence": confidence,
                          "notes": notes, "pay_idx": pidx, "group": ""}

    # Prefer exact Macron/Paymeter reference evidence + VPS virtual account/gross/time.
    direct_edges = []
    direct_candidates_m: dict[int, list[tuple[float, int]]] = defaultdict(list)
    direct_candidates_v: dict[int, list[tuple[float, int]]] = defaultdict(list)
    for mi, mm in enumerate(m):
        pi = mm["_pidx"]
        if pi is None or mm["_identity"] == "CONFLICT" or not mm["_va"]:
            continue
        pp = p[pi]
        if pp["_va"] != mm["_va"] or pp["_gross"] is None or pp["_date"] is None:
            continue
        for vi in by_va_amount.get((mm["_va"], pp["_gross"]), []):
            vv = v[vi]
            if vv["_date"] is None:
                continue
            minutes = abs((vv["_date"] - pp["_date"]).total_seconds()) / 60
            if minutes > max_minutes:
                continue
            expected = (mm["_amount"] + fee if mm["_activity"] == "TOKEN PURCHASE"
                        else mm["_amount"])
            money_ok = expected == vv["_gross"]
            token_ok = mm["_amount"] == pp["_token"]
            score = minutes + (2000 if not money_ok else 0) + (200 if not token_ok else 0)
            direct_edges.append((score, mi, vi, pi, money_ok, token_ok, minutes))
            direct_candidates_m[mi].append((score, vi))
            direct_candidates_v[vi].append((score, mi))
    for score, mi, vi, pi, money_ok, token_ok, gap in sorted(direct_edges):
        if mi in matched_m or vi in matched_v:
            continue
        mm, vv = m[mi], v[vi]
        alt_m = sum(abs(cost - score) <= 10 for cost, other in direct_candidates_m[mi] if other != vi)
        alt_v = sum(abs(cost - score) <= 10 for cost, other in direct_candidates_v[vi] if other != mi)
        unique = alt_m == 0 and alt_v == 0
        high = (gap <= high_confidence_minutes and money_ok and token_ok and unique
                and mm["_identity"] == "KYC")
        issues = []
        if not money_ok:
            issues.append("VPS gross disagrees with Macron fee rule")
        if not token_ok:
            issues.append("Paymeter token disagrees with Macron")
        if not unique:
            issues.append("Alternative transaction candidate within 10 minutes")
        if mm["_identity"] != "KYC":
            issues.append("Macron to virtual account inferred from Paymeter, not verified KYC")
        if gap > high_confidence_minutes:
            issues.append("Paymeter-to-VPS delay exceeds high confidence window")
        assign(mi, vi, "REFERENCE_PAYMETER_VPS", "HIGH" if high else "REVIEW", "; ".join(issues), pi)

    # KYC-to-VPS matching independent of Paymeter; fee must agree exactly.
    kyc_edges = []
    kyc_candidates_m: dict[int, list[tuple[float, int]]] = defaultdict(list)
    kyc_candidates_v: dict[int, list[tuple[float, int]]] = defaultdict(list)
    for mi, mm in enumerate(m):
        if (mi in matched_m or not mm["_va"] or mm["_identity"] == "CONFLICT"
                or mm["_amount"] is None or mm["_date"] is None):
            continue
        expected = [(mm["_amount"] + fee, "TOKEN_FEE100")]
        if mm["_activity"] == "PAYBILL":
            expected = [(mm["_amount"], "PAYBILL_SAME_GROSS"),
                        (mm["_amount"] + fee, "PAYBILL_FEE100")]
        for gross, rule in expected:
            for vi in by_va_amount.get((mm["_va"], gross), []):
                if vi in matched_v:
                    continue
                vv = v[vi]
                if vv["_date"] is None:
                    continue
                minutes = (mm["_date"] - vv["_date"]).total_seconds() / 60
                if abs(minutes) > max_minutes:
                    continue
                day_gap = abs((mm["_date"].date() - vv["_date"].date()).days)
                if day_gap > 1:
                    continue
                # Prefer time proximity, receipts before tokens, and same-day transaction pairs.
                score = abs(minutes) + (250 if minutes < -3 else 0) + (500 if day_gap else 0)
                score += 100 if rule == "PAYBILL_FEE100" else 0
                kyc_edges.append((score, mi, vi, rule, minutes, day_gap))
                kyc_candidates_m[mi].append((score, vi))
                kyc_candidates_v[vi].append((score, mi))
    for score, mi, vi, rule, minutes, day_gap in sorted(kyc_edges):
        if mi in matched_m or vi in matched_v:
            continue
        mm = m[mi]
        unique = (not any(abs(s - score) <= 10 and other != vi for s, other in kyc_candidates_m[mi])
                  and not any(abs(s - score) <= 10 and other != mi for s, other in kyc_candidates_v[vi]))
        high = (mm["_identity"] == "KYC" and rule == "TOKEN_FEE100"
                and day_gap == 0 and -3 <= minutes <= high_confidence_minutes and unique)
        issues = []
        if mm["_identity"] != "KYC":
            issues.append("No independently verified KYC link")
        if not unique:
            issues.append("Multiple plausible account/amount/time pairings")
        if day_gap:
            issues.append("Receipt and token are on different local dates")
        if minutes < -3:
            issues.append("Token timestamp precedes receipt")
        if abs(minutes) > high_confidence_minutes:
            issues.append("Timestamp gap exceeds high confidence window")
        if rule.startswith("PAYBILL"):
            issues.append("PayBill value rule needs separate confirmation")
        assign(mi, vi, "ACCOUNT_AMOUNT_DATE_" + rule,
               "HIGH" if high else "REVIEW", "; ".join(issues), mm["_pidx"])

    # A single VPS PayBill deposit may cover multiple Macron PayBill lines.
    grouped_members: dict[int, str] = {}
    groups = []
    unpaid_paybill: dict[tuple[str, date], list[int]] = defaultdict(list)
    for mi, mm in enumerate(m):
        if (mi not in matched_m and mm["_activity"] == "PAYBILL" and mm["_va"]
                and mm["_date"] and mm["_amount"] is not None):
            unpaid_paybill[(mm["_va"], mm["_date"].date())].append(mi)
    for (va, day), mac_ids in sorted(unpaid_paybill.items()):
        if len(mac_ids) < 2:
            continue
        total = sum(m[i]["_amount"] for i in mac_ids)
        options = [j for j in by_va.get(va, []) if j not in matched_v and v[j]["_date"]
                   and abs((v[j]["_date"].date() - day).days) <= 1
                   and v[j]["_gross"] in (total, total + fee)]
        if len(options) != 1:
            continue
        vi = options[0]
        group_id = f"PAYBILL-{len(groups) + 1:03d}"
        leader = mac_ids[0]
        assign(leader, vi, "PAYBILL_GROUP_AGGREGATE", "REVIEW",
               "One VPS deposit covers several PayBill entries; allocate bank receipt once", m[leader]["_pidx"])
        match_meta[leader]["group"] = group_id
        for mi in mac_ids[1:]:
            grouped_members[mi] = group_id
        groups.append({"group_id": group_id, "virtual_account": va, "date": day,
                       "VPS_source_row": v[vi]["_row"],
                       "Macron_source_rows": ", ".join(str(m[i]["_row"]) for i in mac_ids),
                       "macron_lines": len(mac_ids), "macron_total": ngn(total),
                       "vps_gross": ngn(v[vi]["_gross"])})

    # Potential duplicate tokens: same virtual account + same day + same amount,
    # with more Macron token lines than VPS receipts matching the fee rule.
    tokens_by = defaultdict(list)
    payments_by = defaultdict(list)
    for mi, mm in enumerate(m):
        if mm["_activity"] == "TOKEN PURCHASE" and mm["_va"] and mm["_date"] and mm["_amount"] is not None:
            tokens_by[(mm["_va"], mm["_date"].date(), mm["_amount"])].append(mi)
    for vi, vv in enumerate(v):
        if vv["_va"] and vv["_date"] and vv["_gross"] is not None:
            payments_by[(vv["_va"], vv["_date"].date(), vv["_gross"] - fee)].append(vi)
    repeat_candidates = set()
    for key, ids in tokens_by.items():
        if len(ids) > 1 and len(ids) > len(payments_by.get(key, [])):
            repeat_candidates.update(i for i in ids if i not in matched_m)

    # As an audit convenience, find Paymeter reports linked to VPS-only payments.
    paymeter_by_va_gross = defaultdict(list)
    for pi, pp in enumerate(p):
        if pp["_va"] and pp["_gross"] is not None:
            paymeter_by_va_gross[(pp["_va"], pp["_gross"])].append(pi)

    def paymeter_for_vps_only(vi: int) -> int | None:
        vv = v[vi]
        possibilities = [(abs((p[pi]["_date"] - vv["_date"]).total_seconds()), pi)
                         for pi in paymeter_by_va_gross.get((vv["_va"], vv["_gross"]), [])
                         if vv["_date"] and p[pi]["_date"]
                         and abs((p[pi]["_date"] - vv["_date"]).total_seconds()) <= max_minutes * 60]
        return min(possibilities)[1] if possibilities else None

    posting_rows = []
    kyc_issues = []
    counts = Counter()
    used_vps = set()

    def build_row(mi: int | None, vi: int | None) -> dict:
        mm = m[mi] if mi is not None else None
        vv = v[vi] if vi is not None else None
        if vi is not None:
            if vi in used_vps:
                raise AssertionError("Repeated VPS source row in output")
            used_vps.add(vi)
        va = (vv["_va"] if vv else "") or (mm["_va"] if mm else "")
        lookup = k_va.get(va)
        pi = mm["_pidx"] if mm is not None else paymeter_for_vps_only(vi)
        pp = p[pi] if pi is not None else None
        identity = "KYC_MAPPED" if lookup and lookup["mac"] else "KYC_NAME_ONLY" if lookup else "KYC_NOT_FOUND"
        if mm is not None and lookup and lookup["mac"] and lookup["mac"] != mm["_mac"]:
            identity = "KYC_ACCOUNT_CONFLICT"
        if mm is not None and mm["_identity"] == "CONFLICT":
            identity = "PAYMETER_KYC_REFERENCE_CONFLICT"
        trx_date = (vv["_date"].date() if vv and vv["_date"] else
                    mm["_date"].date() if mm and mm["_date"] else None)
        if mm and vv:
            meta = match_meta[mi]
            status = "MATCHED_HIGH" if meta["confidence"] == "HIGH" else "MATCHED_REVIEW"
            conf, method, notes, group = meta["confidence"], meta["method"], meta["notes"], meta["group"]
        elif mm and mi in grouped_members:
            status, conf, method, notes = ("PAYBILL_GROUP_MEMBER", "REVIEW", "PAYBILL_GROUP",
                                             "Covered by grouped VPS receipt; VPS columns intentionally blank")
            group = grouped_members[mi]
        elif mm:
            status = "MACRON_ONLY_TOKEN" if mm["_activity"] == "TOKEN PURCHASE" else "MACRON_ONLY_PAYBILL"
            conf, method, notes, group = "UNMATCHED", "NO_PAYMENT_PAIR", "", ""
        else:
            status = "VPS_ONLY_PAYMENT"
            conf, method, notes, group = "UNMATCHED", "NO_TOKEN_PAIR", "", ""
        if mi in repeat_candidates:
            notes += ("; " if notes else "") + "Possible repeated token issuance; investigate"
        if vv and vv["_gross"] is not None and vv["_gross"] <= 1000 * 100 and not mm:
            notes += ("; " if notes else "") + "Unmatched small payment (<= NGN 1,000)"
        gross_c = vv["_gross"] if vv else None
        mac_c = mm["_amount"] if mm else None
        sett_c = vv["_settled"] if vv else None
        actual_c = gross_c - sett_c if gross_c is not None and sett_c is not None else None
        reported_c = vv["_charge"] if vv else None
        expected_c = gross_c - fee if gross_c is not None else None
        if expected_c is not None and expected_c < 0:
            notes += ("; " if notes else "") + "Payment below the flat charge; negative net value"
        variance_c = mac_c - expected_c if mac_c is not None and expected_c is not None else None
        if (mm and vv and mm["_activity"] == "TOKEN PURCHASE" and variance_c != 0):
            notes += ("; " if notes else "") + "Token differs from VPS gross minus fee"
            status, conf = "MATCHED_REVIEW", "REVIEW"
        if mm and vv and identity in ("KYC_ACCOUNT_CONFLICT", "PAYMETER_KYC_REFERENCE_CONFLICT"):
            status, conf = "MATCHED_REVIEW", "REVIEW"
            notes += ("; " if notes else "") + "Identity conflict"
        if actual_c is not None and reported_c is not None and actual_c != reported_c:
            notes += ("; " if notes else "") + "VPS charge disagrees with gross minus settlement"
        bank_actual = ngn(actual_c)
        # Keep fields from a missing side *blank*; only posting fields use zero.
        row = {
            "transaction_date": trx_date,
            "customer_name": lookup["name"] if lookup else "",
            "account_number": lookup["va"] if lookup else va,
            "cash_in": ngn(expected_c) if vv and expected_c is not None else 0,
            "cash_out": ngn(mac_c) if mm and mac_c is not None else 0,
            "interest": ngn(fee - actual_c) if vv and actual_c is not None else 0,
            "source": "Bank/Macron" if mm and vv else "Macron" if mm else "Bank",
            "status": status, "confidence": conf, "matching_method": method,
            "review_notes": notes, "duplicate_token_candidate": "YES" if mi in repeat_candidates else "NO",
            "identity_status": identity,
            "Macron_account_KYC": lookup["mac"] if lookup else "",
            "virtual_account_KYC": lookup["va"] if lookup else "",
            "Macron_source_row": mm["_row"] if mm else None,
            "VPS_source_row": vv["_row"] if vv else None,
            "Paymeter_source_row": pp["_row"] if pp else None,
            "Paymeter_reference": pp.get("Reference", "") if pp else "",
            "Paymeter_input_amount": ngn(pp["_gross"]) if pp else None,
            "Paymeter_token_amount": ngn(pp["_token"]) if pp else None,
            "PayBill_group": group, "VPS_gross_NGN": ngn(gross_c),
            "VPS_actual_bank_charge_NGN": bank_actual,
            "VPS_reported_bank_charge_NGN": ngn(reported_c),
            "VPS_charge_difference_NGN": ngn(reported_c - actual_c)
            if reported_c is not None and actual_c is not None else None,
            "VPS_settled_NGN": ngn(sett_c), "expected_token_NGN": ngn(expected_c),
            "token_variance_NGN": ngn(variance_c),
        }
        for field in MACRON_FIELDS:
            row["Macron | " + field] = mm.get(field) if mm else None
        for field in VPS_FIELDS:
            row["VPS | " + field] = vv.get(field) if vv else None
        counts[status] += 1
        if identity != "KYC_MAPPED":
            kyc_issues.append(row)
        return row

    for mi in range(len(m)):
        posting_rows.append(build_row(mi, matched_m.get(mi)))
    for vi in range(len(v)):
        if vi not in matched_v:
            posting_rows.append(build_row(None, vi))

    assert len(used_vps) == len(v), "Not all bank records are in report"
    assert len(posting_rows) == len(m) + len(v) - len(matched_m)
    assert sum(r["Macron_source_row"] is not None for r in posting_rows) == len(m)
    assert sum(r["VPS_source_row"] is not None for r in posting_rows) == len(v)
    assert len({r["VPS_source_row"] for r in posting_rows if r["VPS_source_row"] is not None}) == len(v)
    assert len({r["Macron_source_row"] for r in posting_rows if r["Macron_source_row"] is not None}) == len(m)

    customer_rows = defaultdict(lambda: {"customer_name": "", "virtual_account": "", "macron_account": "",
                                         "cash_in": 0.0, "cash_out": 0.0, "interest": 0.0,
                                         "matched": 0, "vps_only": 0, "macron_only": 0})
    for row in posting_rows:
        account = row["account_number"] or "UNMAPPED"
        entry = customer_rows[account]
        entry["virtual_account"] = account
        entry["customer_name"] = entry["customer_name"] or row["customer_name"]
        entry["macron_account"] = entry["macron_account"] or row["Macron_account_KYC"]
        for col in ("cash_in", "cash_out", "interest"):
            entry[col] += float(row[col])
        if row["status"].startswith("MATCHED_"):
            entry["matched"] += 1
        elif row["status"].startswith("VPS_ONLY"):
            entry["vps_only"] += 1
        elif row["status"].startswith("MACRON_ONLY"):
            entry["macron_only"] += 1
    exposure = []
    for entry in customer_rows.values():
        entry = dict(entry)
        entry["net_cash_in_less_cash_out"] = round(entry["cash_in"] - entry["cash_out"], 2)
        for c in ("cash_in", "cash_out", "interest"):
            entry[c] = round(entry[c], 2)
        exposure.append(entry)
    exposure.sort(key=lambda e: abs(e["net_cash_in_less_cash_out"]), reverse=True)

    mapped = [{"virtual_account": obj["va"], "macron_account": obj["mac"],
               "customer_name": obj["name"], "meter_number": obj["meter"],
               "KYC_source_row": obj["row"],
               "status": "MAPPED" if obj["mac"] else "NAME_ONLY"}
              for obj in k_va.values()]
    missing_small = [row for row in posting_rows if row["status"] == "VPS_ONLY_PAYMENT"
                     and (as_cents(row["VPS_gross_NGN"]) or 0) <= 1000 * 100]
    duplicate_rows = [row for row in posting_rows if row["duplicate_token_candidate"] == "YES"]
    used_paymeter_refs = {mm["_pidx"] for mm in m if mm["_pidx"] is not None}
    paymeter_only = [pp for pi, pp in enumerate(p) if pi not in used_paymeter_refs]
    return {"rows": posting_rows, "summary": counts, "customer": exposure, "kyc": mapped,
            "kyc_issues": kyc_issues, "groups": groups, "small": missing_small,
            "duplicate": duplicate_rows, "paymeter_only": paymeter_only,
            "macron_total": len(m), "vps_total": len(v), "paymeter_total": len(p),
            "kyc_total": len(k), "matched_count": len(matched_m),
            "full_mappings": sum(bool(x["macron_account"]) for x in mapped),
            "source_inputs": data, "fee_ngn": fee_ngn, "max_delay_hours": max_delay_hours,
            "high_confidence_minutes": high_confidence_minutes}


def excel_bytes(report: dict) -> bytes:
    """Produce a formatted workbook; treat all source strings as literal strings."""
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {"in_memory": True, "strings_to_formulas": False,
                                    "strings_to_urls": False, "nan_inf_to_errors": True})
    normal = wb.add_format({"font_name": "Aptos", "font_size": 10})
    txt = wb.add_format({"font_name": "Aptos", "font_size": 10, "num_format": "@"})
    money = wb.add_format({"font_name": "Aptos", "num_format": '#,##0.00;[Red](#,##0.00);0.00'})
    count_fmt = wb.add_format({"num_format": "0"})
    dates = wb.add_format({"font_name": "Aptos", "num_format": "dd-mmm-yy"})
    head = wb.add_format({"bold": True, "bg_color": "#14344C", "font_color": "#FFFFFF",
                          "text_wrap": True, "valign": "vcenter", "border": 0})
    posting_head = wb.add_format({"bold": True, "bg_color": "#077F75", "font_color": "#FFFFFF",
                                  "text_wrap": True, "valign": "vcenter"})
    warn = wb.add_format({"bg_color": "#FFF0DB", "font_color": "#814600"})
    red = wb.add_format({"bg_color": "#FCE5E5", "font_color": "#972D27"})
    green = wb.add_format({"bg_color": "#E2F5EB", "font_color": "#146B42"})
    row_heads = POSTING + AUDIT + ["Macron | " + f for f in MACRON_FIELDS] + ["VPS | " + f for f in VPS_FIELDS]
    monetary = {"cash_in", "cash_out", "interest", "VPS_gross_NGN", "VPS_actual_bank_charge_NGN",
                "VPS_reported_bank_charge_NGN", "VPS_charge_difference_NGN", "VPS_settled_NGN",
                "expected_token_NGN", "token_variance_NGN", "Paymeter_input_amount", "Paymeter_token_amount",
                "net_cash_in_less_cash_out", "macron_total", "vps_gross", "Input Amount", "Transaction Amount",
                "Macron | AMOUNT", "VPS | transaction_amount_minor", "VPS | settled_amount_minor",
                "VPS | charge_amount_minor"}
    # A few raw source fields are monetary; all other identifiers are written as text.
    id_number_fields = {"Macron_source_row", "VPS_source_row", "Paymeter_source_row"}

    def write_value(ws, r, c, val, name):
        if val is None or val == "":
            return
        if isinstance(val, date):
            dt = datetime.combine(val, time.min) if not isinstance(val, datetime) else val
            ws.write_datetime(r, c, dt, dates)
        elif name in monetary:
            cents = as_cents(val)
            if cents is not None:
                ws.write_number(r, c, ngn(cents), money)
            else:
                ws.write_string(r, c, norm(val), txt)
        elif name in id_number_fields or (isinstance(val, int) and name in ("matched", "macron_only", "vps_only", "macron_lines", "KYC_source_row", "Value")):
            ws.write_number(r, c, int(val), count_fmt)
        elif name == "Value" and isinstance(val, float):
            ws.write_number(r, c, val, money)
        else:
            # Explicit write_string also prevents formula injection from uploaded names and narrations.
            ws.write_string(r, c, norm(val)[:32767], txt)

    def make_sheet(name: str, headers: list[str], rows: list[dict], filter_rows=True) -> None:
        ws = wb.add_worksheet(name[:31])
        ws.hide_gridlines(2)
        ws.set_default_row(18)
        ws.set_row(0, 34)
        ws.freeze_panes(1, 7 if name == "Combined" else 0)
        for c, title in enumerate(headers):
            ws.write_string(0, c, title, posting_head if name == "Combined" and c < 7 else head)
            width = 18
            if title in ("customer_name", "source_acct_name", "VPS | source_acct_name"):
                width = 31
            if title in ("review_notes", "matching_method", "VPS | narration"):
                width = 43
            if any(word in title.lower() for word in ("reference", "transaction id", "session_id")):
                width = 39
            if title in ("status", "identity_status"):
                width = 28
            if title == "transaction_date":
                width = 17
            ws.set_column(c, c, min(width, 55))
        for r, item in enumerate(rows, 1):
            for c, title in enumerate(headers):
                write_value(ws, r, c, item.get(title), title)
        if filter_rows and rows:
            ws.autofilter(0, 0, len(rows), len(headers) - 1)
        if name == "Combined" and rows:
            ci = headers.index("status")
            col = xlsxwriter.utility.xl_col_to_name(ci)
            area = (1, ci, len(rows), ci)
            ws.conditional_format(*area, {"type": "text", "criteria": "containing", "value": "VPS_ONLY", "format": warn})
            ws.conditional_format(*area, {"type": "text", "criteria": "containing", "value": "MACRON_ONLY", "format": red})
            ws.conditional_format(*area, {"type": "text", "criteria": "containing", "value": "MATCHED_HIGH", "format": green})

    combined = report["rows"]
    make_sheet("Combined", row_heads, combined)
    matches = [r for r in combined if r["status"] in ("MATCHED_HIGH", "MATCHED_REVIEW")]
    macros = [r for r in combined if r["status"] in ("MACRON_ONLY_TOKEN", "MACRON_ONLY_PAYBILL")]
    vpsonly = [r for r in combined if r["status"] == "VPS_ONLY_PAYMENT"]
    make_sheet("VPS No Token", row_heads, vpsonly)
    make_sheet("Macron No Payment", row_heads, macros)
    make_sheet("Review Matches", row_heads, [r for r in matches if r["confidence"] == "REVIEW"])
    make_sheet("Small Unmatched", row_heads, report["small"])
    make_sheet("Possible Duplicates", row_heads, report["duplicate"])
    make_sheet("KYC Mapping", ["virtual_account", "macron_account", "customer_name", "meter_number",
                               "KYC_source_row", "status"], report["kyc"])
    customer_headers = ["virtual_account", "macron_account", "customer_name", "cash_in", "cash_out",
                        "interest", "matched", "vps_only", "macron_only", "net_cash_in_less_cash_out"]
    make_sheet("Customer Exposure", customer_headers, report["customer"])
    make_sheet("PayBill Groups", ["group_id", "virtual_account", "date", "VPS_source_row",
                                  "Macron_source_rows", "macron_lines", "macron_total", "vps_gross"], report["groups"])
    pay_only = [{"Paymeter_source_row": x["_row"], "Reference": x.get("Reference"),
                 "Account Number": x.get("Account Number"), "Created At": x.get("Created At"),
                 "Input Amount": x.get("Input Amount"), "Transaction Amount": x.get("Transaction Amount")}
                for x in report["paymeter_only"]]
    make_sheet("Paymeter No Reference", ["Paymeter_source_row", "Reference", "Account Number",
                                           "Created At", "Input Amount", "Transaction Amount"], pay_only)

    summary = [
        {"Metric": "Macron transactions", "Value": report["macron_total"]},
        {"Metric": "VPS bank receipts", "Value": report["vps_total"]},
        {"Metric": "Paymeter records (supporting only)", "Value": report["paymeter_total"]},
        {"Metric": "KYC records", "Value": report["kyc_total"]},
        {"Metric": "KYC records with validated Macron account", "Value": report["full_mappings"]},
        {"Metric": "Matched one-to-one pairs", "Value": report["matched_count"]},
        {"Metric": "MATCHED_HIGH", "Value": report["summary"]["MATCHED_HIGH"]},
        {"Metric": "MATCHED_REVIEW", "Value": report["summary"]["MATCHED_REVIEW"]},
        {"Metric": "MACRON_ONLY_TOKEN", "Value": report["summary"]["MACRON_ONLY_TOKEN"]},
        {"Metric": "MACRON_ONLY_PAYBILL", "Value": report["summary"]["MACRON_ONLY_PAYBILL"]},
        {"Metric": "VPS_ONLY_PAYMENT", "Value": report["summary"]["VPS_ONLY_PAYMENT"]},
        {"Metric": "PAYBILL_GROUP_MEMBER", "Value": report["summary"]["PAYBILL_GROUP_MEMBER"]},
        {"Metric": "Small VPS-only receipts <= NGN 1000", "Value": len(report["small"])},
        {"Metric": "Possible duplicate token rows", "Value": len(report["duplicate"])},
        {"Metric": "Total combined rows", "Value": len(combined)},
        {"Metric": "Unmatched VPS gross (NGN)", "Value": round(sum(r["VPS_gross_NGN"] or 0 for r in vpsonly), 2)},
        {"Metric": "Unmatched Macron tokens (NGN)", "Value": round(sum(r["cash_out"] for r in macros if r["status"] == "MACRON_ONLY_TOKEN"), 2)},
        {"Metric": "Charge per transaction (NGN)", "Value": report["fee_ngn"]},
        {"Metric": "Maximum pairing time window (hours)", "Value": report["max_delay_hours"]},
    ]
    make_sheet("Summary", ["Metric", "Value"], summary)
    rules = [
        {"Rule": "Ledgers", "Details": "VPS = payments received. Macron = tokens/PayBill issued. Paymeter supports reference matching."},
        {"Rule": "Identity", "Details": "VPS virtual_acct_no = KYC Paymeter account; KYC maps virtual account to Macron ACCOUNT NUMBER."},
        {"Rule": "Reference", "Details": "Macron REFERENCE ID = Paymeter Reference; Paymeter Account Number = VPS virtual_acct_no."},
        {"Rule": "Amounts", "Details": "Token purchases: VPS gross = Macron amount + fixed NGN fee. PayBill may use same gross or grouped allocations, flagged for review."},
        {"Rule": "Time", "Details": "Use VPS cba_tran_time in Nigeria local time, otherwise UTC created_at converted to Lagos +01:00. Macron time local."},
        {"Rule": "Pairing", "Details": "No VPS bank receipt or Macron transaction is paired more than once. Near-ties and delays marked REVIEW."},
        {"Rule": "Unmatched", "Details": "Macron/VPS side stays blank when missing. Only the seven posting financial fields show zero for absent sides."},
        {"Rule": "Posting date", "Details": "VPS date first, Macron date fallback; saved as real date dd-mmm-yy, with no timestamp."},
        {"Rule": "cash_in", "Details": "VPS gross - fixed NGN fee; negative values for gross < fee are visible and flagged."},
        {"Rule": "cash_out", "Details": "Actual Macron amount. PayBill is not automatically considered a prepaid token."},
        {"Rule": "interest", "Details": "Fixed NGN fee - actual bank charge; actual bank charge = VPS gross - VPS settled."},
        {"Rule": "VPS _minor", "Details": "Values treated as NGN as observed in provided VPS sample (e.g., 5,000 represents NGN 5,000, NOT 50)."},
        {"Rule": "Fee discrepancy", "Details": "VPS reported charge - (VPS gross - VPS settled); any discrepancy flagged."},
        {"Rule": "Status", "Details": "MATCHED_REVIEW is provisional; never post a customer debt solely because of a screening flag."},
        {"Rule": "PayBill group", "Details": "One VPS bank record appears on one group-leader row only; member rows remain with blank VPS columns and group identifier."},
        {"Rule": "Security", "Details": "Uploaded names and narration exported as literal strings, never active Excel formulas. Files processed in RAM."},
    ]
    make_sheet("Methodology", ["Rule", "Details"], rules)
    wb.close()
    output.seek(0)
    return output.getvalue()


def main() -> None:
    import streamlit as st

    st.set_page_config(page_title="Macron - VPS Reconciliation", layout="wide")
    st.title("Macron ↔ VPS Financial Reconciliation")
    st.caption("Primary ledgers: VPS payments and Macron tokens. Paymeter references and KYC provide identity evidence.")
    st.info("Transactions are processed in memory. For sensitive customer information, run this locally or on a secured private server.")
    with st.sidebar:
        st.header("Matching rules")
        fee = st.number_input("Fixed customer charge (NGN)", min_value=0, max_value=5000, value=100, step=10)
        hours = st.slider("Maximum transaction delay (hours)", min_value=1, max_value=72, value=36)
        confidence = st.slider("High-confidence window (minutes)", min_value=5, max_value=120, value=30)
        st.caption("Uncertain matches stay flagged for review. Bank receipts are never reused.")
    a, b = st.columns(2)
    with a:
        m_file = st.file_uploader("1. Macron transaction Excel", type=["xlsx"], key="m")
        p_file = st.file_uploader("2. Paymeter transactions CSV", type=["csv"], key="p")
    with b:
        v_file = st.file_uploader("3. VPS bank Excel", type=["xlsx"], key="v")
        k_file = st.file_uploader("4. KYC mapping Excel", type=["xlsx"], key="k")
    available = all([m_file, p_file, v_file, k_file])
    if not available:
        st.write("Upload the four source files, then click **Generate reconciliation**.")
        return
    if st.button("Generate reconciliation", type="primary", use_container_width=True):
        try:
            with st.spinner("Reading, linking, auditing and exporting transactions..."):
                raw_inputs = [f.getvalue() for f in (m_file, p_file, v_file, k_file)]
                if any(len(x) > MAX_FILE_BYTES for x in raw_inputs):
                    raise ValueError("Individual file size must not exceed 30 MB")
                inputs = prepare_inputs(raw_inputs[0], raw_inputs[1], raw_inputs[2], raw_inputs[3])
                report = reconcile(inputs, int(fee), hours, confidence)
                binary = excel_bytes(report)
                st.session_state["recon_result"] = report
                st.session_state["recon_binary"] = binary
        except (ValueError, AssertionError, zipfile.BadZipFile, ET.ParseError) as exc:
            st.error(f"Cannot safely complete reconciliation: {exc}")
            return
        except Exception:
            st.error("An unexpected error occurred. Check the files and consult the application logs.")
            raise
    if "recon_result" not in st.session_state:
        return
    report = st.session_state["recon_result"]
    counts = report["summary"]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Matched high confidence", counts["MATCHED_HIGH"])
    c2.metric("Matched: review required", counts["MATCHED_REVIEW"])
    c3.metric("Bank payments with no token", counts["VPS_ONLY_PAYMENT"])
    c4.metric("Tokens without paired payment", counts["MACRON_ONLY_TOKEN"])
    st.download_button(
        "Download complete Excel reconciliation",
        data=st.session_state["recon_binary"],
        file_name="Macron_VPS_Paymeter_KYC_Reconciliation.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary", use_container_width=True,
    )
    st.warning("Unmatched and duplicate statuses indicate items to investigate, not proven financial liabilities."
               " PayBill entries and review matches must be checked before posting.")
    with st.expander("Preview transactions (first 200 rows)", expanded=True):
        from itertools import islice
        rows = report["rows"][:200]
        preview_cols = POSTING + ["status", "review_notes", "Macron_source_row", "VPS_source_row",
                                  "VPS_gross_NGN", "VPS_actual_bank_charge_NGN", "token_variance_NGN"]
        st.dataframe([{k: (r[k].strftime("%d-%b-%y") if isinstance(r[k], date) else r[k])
                       for k in preview_cols} for r in rows], use_container_width=True, hide_index=True)
    with st.expander("Customer exposure (first 100 accounts)"):
        st.dataframe(report["customer"][:100], use_container_width=True, hide_index=True)
    st.caption("The downloaded workbook includes a full Combined sheet, KYC mapping, unmatched lists, "
               "PayBill groups, possible duplicate tokens, accounting fields, and methodology.")


if __name__ == "__main__":
    main()
