import streamlit as st
import pandas as pd
import pdfplumber
import base64


APP_PASSWORD = "admin123"

password = st.text_input("🔒 Enter Password", type="password")
if password:
    if password != APP_PASSWORD:
        st.error("❌ Wrong password")
        st.stop()
else:
    st.info("Please enter password")
    st.stop()

st.title("Ledger Analyzer — Reconciliation Statement")

col_up1, col_up2 = st.columns(2)
with col_up1:
    file1 = st.file_uploader(
        "📁 Upload Company Books", type=["csv", "xlsx", "xls", "pdf"]
    )
with col_up2:
    file2 = st.file_uploader(
        "📁 Upload Party Statement", type=["csv", "xlsx", "xls", "pdf"]
    )

RECON_GROUPS = [
    {
        "label": "Opening Balance",
        "add_less": "Add",
        "co_search": ["opening balance"],
        "party_search": ["opening balance"],
        "search_field": "ledger",
        "exclude": [],
    },
    {
        "label": "Purchase",
        "add_less": "Add",
        "co_search": ["purchase trade", "purchase"],
        "party_search": ["sales", "sale"],
        "search_field": "vch_type",
        "exclude": [],
    },
    {
        "label": "Debit Note-QCR / GRN",
        "add_less": "Less",
        "co_search": ["debit note-qcr", "debit note-grn"],
        "party_search": ["debit note", "credit note"],
        "search_field": "both",
        "exclude": [],
    },
    {
        "label": "TDS",
        "add_less": "Less",
        "co_search": ["tds", "tax deducted"],
        "party_search": ["tds", "tax deducted"],
        "search_field": "both",
        "exclude": [],
    },
    {
        "label": "Payment / Receipt / Bank / Journal",
        "add_less": "Add",
        "co_search": ["payment", "receipt", "bank", "journal"],
        "party_search": [
            "payment recd",
            "payment received",
            "bank",
            "receipt",
            "journal",
        ],
        "search_field": "both",
        "exclude": ["tds", "tax deducted"],
    },
]


# ──────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────
def detect_header(df_raw):
    for i in range(len(df_raw)):
        row = df_raw.iloc[i].fillna("").astype(str).str.lower().tolist()
        if (
            any("particulars" in c for c in row)
            and any("debit" in c or c.strip() == "dr" for c in row)
            and any("credit" in c or c.strip() == "cr" for c in row)
        ):
            return i
    return None


def detect_col_positions(df_raw, header_row):
    header = df_raw.iloc[header_row].fillna("").astype(str).str.lower().tolist()
    pos = {
        "date": None,
        "particulars": None,
        "ledger": None,
        "vch_type": None,
        "vch_no": None,
        "debit": None,
        "credit": None,
    }
    for i, cell in enumerate(header):
        cl = cell.strip()
        if "date" in cl and pos["date"] is None:
            pos["date"] = i
        elif "particulars" in cl and pos["particulars"] is None:
            pos["particulars"] = i
            pos["ledger"] = i + 1
        elif ("vch type" in cl or "voucher type" in cl) and pos["vch_type"] is None:
            pos["vch_type"] = i
        elif ("vch no" in cl or "voucher no" in cl) and pos["vch_no"] is None:
            pos["vch_no"] = i
        elif cl in ["debit", "dr"] and pos["debit"] is None:
            pos["debit"] = i
        elif cl in ["credit", "cr"] and pos["credit"] is None:
            pos["credit"] = i
    return pos


def clean_date(val):
    try:
        return pd.to_datetime(val).strftime("%d-%m-%Y")
    except:
        return str(val).strip().strip("'")


def parse_amount(val):
    try:
        return round(float(str(val).replace(",", "").strip()), 2)
    except:
        return 0.0


def safe_cell(row, idx):
    if idx >= len(row):
        return ""
    v = row[idx]
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("none", "nan") else s


def is_to_by_pdf(rows):
    valid = [r for r in rows if r and any(c for c in r if c)]
    if not valid:
        return False
    s = [safe_cell(r, 1) for r in valid[:30]]
    return sum(1 for v in s if v in ["To", "By", "Dr", "Cr"]) >= 3


def is_two_col_pdf(rows):
    valid = [r for r in rows if r and any(c for c in r if c)]
    if not valid:
        return False
    avg = sum(len(r) for r in valid[:20]) / len(valid[:20])
    s = [safe_cell(r, 1) for r in valid[:10]]
    return avg >= 6 and not any(v in ["To", "By"] for v in s)


def extract_two_col_pdf(all_rows):
    records = []
    for row in all_rows:
        row = list(row) + [None] * 8
        for di, pi, ai, typ in [(0, 1, 3, "debit"), (4, 5, 7, "credit")]:
            d = safe_cell(row, di)
            p = safe_cell(row, pi)
            a = parse_amount(safe_cell(row, ai))
            if d and p and a > 0:
                vno, leg = "", p
                if " - " in p:
                    parts = p.split(" - ", 1)
                    leg, vno = parts[0].strip(), parts[1].strip()
                records.append(
                    {
                        "Date": clean_date(d),
                        "Vch Type": leg,
                        "Vch No.": vno,
                        "Ledger": leg,
                        "Debit": a if typ == "debit" else 0.0,
                        "Credit": a if typ == "credit" else 0.0,
                    }
                )
    if not records:
        return None
    df = pd.DataFrame(records)
    return df[(df["Debit"] != 0) | (df["Credit"] != 0)].reset_index(drop=True)


def extract_to_by_pdf(all_rows):
    records, last_date = [], ""
    skip = {"brought forward", "carried over", "none", "", "nan"}
    TO_VALS = {"To", "By", "Dr", "Cr"}
    for raw_row in all_rows:
        row = list(raw_row) + [None] * 8
        c0, c1, c2, c3, c4, c5, c6 = [safe_cell(row, i) for i in range(7)]
        if c0 and c0 not in TO_VALS:
            try:
                last_date = clean_date(c0)
            except:
                pass
        if c1 in TO_VALS and c2:
            a3, a4 = parse_amount(c3), parse_amount(c4)
            if a3 > 0 and not any(ch.isalpha() for ch in c3.replace(",", "")):
                if c2.lower() not in skip:
                    records.append(
                        {
                            "Date": last_date,
                            "Vch Type": c2,
                            "Vch No.": "",
                            "Ledger": c2,
                            "Debit": a3 if c1 in ["To", "Dr"] else 0.0,
                            "Credit": a3 if c1 in ["By", "Cr"] else 0.0,
                        }
                    )
                continue
            if a4 > 0 and not any(ch.isalpha() for ch in c4.replace(",", "")) and c3:
                part = f"{c2} {c3}".strip()
                if part.lower() not in skip:
                    records.append(
                        {
                            "Date": last_date,
                            "Vch Type": part,
                            "Vch No.": "",
                            "Ledger": part,
                            "Debit": a4 if c1 in ["To", "Dr"] else 0.0,
                            "Credit": a4 if c1 in ["By", "Cr"] else 0.0,
                        }
                    )
                continue
            debit, credit = parse_amount(c5), parse_amount(c6)
            if c2.lower() in skip or (debit == 0 and credit == 0):
                continue
            records.append(
                {
                    "Date": last_date,
                    "Vch Type": c3 if c3 else c2,
                    "Vch No.": c4,
                    "Ledger": c2,
                    "Debit": debit if c1 in ["To", "Dr"] else 0.0,
                    "Credit": credit if c1 in ["By", "Cr"] else 0.0,
                }
            )
            continue
        if c0 in TO_VALS and c1:
            a2, a3 = parse_amount(c2), parse_amount(c3)
            if a2 > 0 and not any(ch.isalpha() for ch in c2.replace(",", "")):
                if c1.lower() not in skip:
                    records.append(
                        {
                            "Date": last_date,
                            "Vch Type": c1,
                            "Vch No.": "",
                            "Ledger": c1,
                            "Debit": a2 if c0 in ["To", "Dr"] else 0.0,
                            "Credit": a2 if c0 in ["By", "Cr"] else 0.0,
                        }
                    )
                continue
            if a3 > 0 and not any(ch.isalpha() for ch in c3.replace(",", "")) and c2:
                part = f"{c1} {c2}".strip()
                if part.lower() not in skip:
                    records.append(
                        {
                            "Date": last_date,
                            "Vch Type": part,
                            "Vch No.": "",
                            "Ledger": part,
                            "Debit": a3 if c0 in ["To", "Dr"] else 0.0,
                            "Credit": a3 if c0 in ["By", "Cr"] else 0.0,
                        }
                    )
                continue
    if not records:
        return None
    df = pd.DataFrame(records)
    return df[(df["Debit"] != 0) | (df["Credit"] != 0)].reset_index(drop=True)


def get_word_rows(page):
    words = page.extract_words()
    if not words:
        return []
    lines = {}
    for w in words:
        y = round(w["top"] / 3) * 3
        lines.setdefault(y, []).append((w["x0"], w["text"]))
    return [[t for _, t in sorted(lines[y])] for y in sorted(lines)]


def read_file(uploaded_file):
    ft = uploaded_file.name.split(".")[-1].lower()
    if ft == "csv":
        return pd.read_csv(uploaded_file, header=None, dtype=str), "standard"
    elif ft in ["xlsx", "xls"]:
        return pd.read_excel(uploaded_file, header=None, dtype=str), "standard"
    elif ft == "pdf":
        all_rows = []
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                word_rows = get_word_rows(page)
                if table:
                    table_flat = " ".join(
                        " ".join(str(c).lower() for c in r if c) for r in table
                    )
                    for wr in word_rows:
                        wr_str = " ".join(wr).lower()
                        if (
                            "opening balance" in wr_str
                            and "opening balance" not in table_flat
                        ):
                            all_rows.append(wr)
                        if (
                            "closing balance" in wr_str
                            and "closing balance" not in table_flat
                        ):
                            all_rows.append(wr)
                    all_rows.extend(table)
                else:
                    all_rows.extend(word_rows)
        if all_rows:
            mc = max(len(r) for r in all_rows)
            df_check = pd.DataFrame(
                [r + [None] * (mc - len(r)) for r in all_rows]
            ).astype(str)
            if detect_header(df_check) is not None:
                return df_check, "standard"
            if is_to_by_pdf(all_rows):
                return all_rows, "to_by"
            elif is_two_col_pdf(all_rows):
                return all_rows, "two_col"
            else:
                return df_check, "standard"
    return None, None


# ──────────────────────────────────────────────────────────────
# CLOSING BALANCE — x-position based detection for PDFs
# ──────────────────────────────────────────────────────────────
def get_debit_credit_x_positions(uploaded_file):
    """Scan PDF for Debit/Credit header x-positions (handles 'Debit Amount' too)."""
    try:
        uploaded_file.seek(0)
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                words = page.extract_words()
                lines = {}
                for w in words:
                    y = round(w["top"] / 3) * 3
                    lines.setdefault(y, []).append(w)
                debit_x, credit_x = None, None
                for y, lw in lines.items():
                    texts = [w["text"].strip().lower() for w in lw]
                    full_line = " ".join(texts)
                    if debit_x is None:
                        if (
                            "debit amount" in full_line
                            or "debit" in texts
                            or "dr" in texts
                        ):
                            for w in lw:
                                if w["text"].strip().lower() in ("debit", "dr"):
                                    debit_x = w["x0"]
                                    break
                    if credit_x is None:
                        if (
                            "credit amount" in full_line
                            or "credit" in texts
                            or "cr" in texts
                        ):
                            for w in lw:
                                if w["text"].strip().lower() in ("credit", "cr"):
                                    credit_x = w["x0"]
                                    break
                if debit_x is not None and credit_x is not None:
                    return debit_x, credit_x
        return None, None
    except:
        return None, None


def get_closing_x_position(uploaded_file):
    """Find x-position of the closing balance amount in PDF."""
    try:
        uploaded_file.seek(0)
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                words = page.extract_words()
                lines = {}
                for w in words:
                    y = round(w["top"] / 3) * 3
                    lines.setdefault(y, []).append(w)
                for y in sorted(lines):
                    lw = lines[y]
                    texts = [w["text"].lower() for w in lw]
                    if any("closing" in t for t in texts) and any(
                        "balance" in t for t in texts
                    ):
                        for w in lw:
                            if parse_amount(w["text"]) > 0:
                                return w["x0"]
        return None
    except:
        return None


def extract_closing_balance(df_raw, file_fmt, cols=None, uploaded_file=None):
    if file_fmt == "standard" and cols is not None:
        for i in range(len(df_raw)):
            row = df_raw.iloc[i].fillna("")
            full_text = " ".join(
                str(row.iloc[j]).strip().lower() for j in range(len(row))
            )
            if "closing balance" not in full_text:
                continue

            debit = (
                parse_amount(row.iloc[cols["debit"]])
                if cols["debit"] is not None and cols["debit"] < len(row)
                else 0
            )
            credit = (
                parse_amount(row.iloc[cols["credit"]])
                if cols["credit"] is not None and cols["credit"] < len(row)
                else 0
            )

            # Trust column position from header — ignore Dr/Cr indicator
            if credit > 0 and debit == 0:
                return credit, "credit"
            if debit > 0 and credit == 0:
                return debit, "debit"
            if debit > 0 and credit > 0:
                return credit, "credit"

            if cols.get("vch_type") is not None and cols["vch_type"] < len(row):
                vt = parse_amount(row.iloc[cols["vch_type"]])
                if vt > 0:
                    return vt, "debit"
        return 0.0, "debit"

    # Non-standard (word rows / to_by / two_col)
    rows = df_raw if isinstance(df_raw, list) else df_raw.values.tolist()
    for raw_row in rows:
        row = list(raw_row) + [None] * 10
        full_text = " ".join(safe_cell(row, i).lower() for i in range(len(row)))
        if "closing balance" not in full_text:
            continue
        amounts = []
        for idx in range(len(row)):
            amt = parse_amount(safe_cell(row, idx))
            if amt > 0:
                amounts.append(amt)
        if not amounts:
            continue
        last_amt = amounts[-1]

        # Use PDF header x-position to determine debit/credit side
        if uploaded_file is not None:
            debit_x = get_debit_credit_x_positions(uploaded_file)[0]
            credit_x = get_debit_credit_x_positions(uploaded_file)[1]
            closing_x = get_closing_x_position(uploaded_file)
            if debit_x is not None and credit_x is not None and closing_x is not None:
                if abs(closing_x - credit_x) <= abs(closing_x - debit_x):
                    return last_amt, "credit"
                else:
                    return last_amt, "debit"

        # Final fallback — default credit
        return last_amt, "credit"
    return 0.0, "debit"


def extract_transactions(df_raw, label, file_fmt):
    if file_fmt == "to_by":
        df = extract_to_by_pdf(df_raw)
        if df is None:
            st.error(f"[{label}] Parse failed.")
            return None, None
        df["Debit"] = df["Debit"].round(2)
        df["Credit"] = df["Credit"].round(2)
        return df, None
    if file_fmt == "two_col":
        df = extract_two_col_pdf(df_raw)
        if df is None:
            st.error(f"[{label}] Parse failed.")
            return None, None
        df["Debit"] = df["Debit"].round(2)
        df["Credit"] = df["Credit"].round(2)
        return df, None
    hr = detect_header(df_raw)
    if hr is None:
        st.error(f"[{label}] Header not found.")
        return None, None
    cols = detect_col_positions(df_raw, hr)
    missing = [
        k
        for k, v in cols.items()
        if v is None and k not in ["vch_type", "vch_no", "date"]
    ]
    if missing:
        st.error(f"[{label}] Missing: {missing}")
        return None, None
    dr = df_raw.iloc[hr + 1 :].reset_index(drop=True)
    date_col = (
        dr.iloc[:, cols["date"]].apply(clean_date)
        if cols["date"] is not None
        else pd.Series([""] * len(dr))
    )
    leg_col = dr.iloc[:, cols["ledger"]].astype(str).str.strip().str.strip("'")
    vt_col = (
        dr.iloc[:, cols["vch_type"]].astype(str).str.strip().str.strip("'")
        if cols["vch_type"] is not None
        else pd.Series([""] * len(dr))
    )
    vn_col = (
        dr.iloc[:, cols["vch_no"]].astype(str).str.strip().str.strip("'")
        if cols["vch_no"] is not None
        else pd.Series([""] * len(dr))
    )
    db_col = (
        pd.to_numeric(dr.iloc[:, cols["debit"]], errors="coerce").fillna(0).round(2)
    )
    cr_col = (
        pd.to_numeric(dr.iloc[:, cols["credit"]], errors="coerce").fillna(0).round(2)
    )
    df = pd.DataFrame(
        {
            "Date": date_col,
            "Vch Type": vt_col,
            "Vch No.": vn_col,
            "Ledger": leg_col,
            "Debit": db_col,
            "Credit": cr_col,
        }
    )
    df = df[(df["Debit"] != 0) | (df["Credit"] != 0)]
    junk = [
        "",
        "nan",
        "none",
        "particulars",
        "vch type",
        "voucher type",
        "debit",
        "credit",
        "total",
        "grand total",
        "closing balance",
    ]
    df = df[~df["Ledger"].str.lower().isin(junk)]
    df = df[
        ~df["Ledger"]
        .str.replace(",", "")
        .str.replace(".", "")
        .str.replace("-", "")
        .str.strip()
        .str.isnumeric()
    ]
    return df.reset_index(drop=True), cols


def search_df(df, keywords, field="both", exclude=None):
    if not keywords:
        return df.iloc[0:0]
    kws = [k.lower() for k in keywords]
    mask = pd.Series([False] * len(df), index=df.index)
    if field in ("vch_type", "both"):
        for k in kws:
            mask |= df["Vch Type"].str.lower().str.strip("'").str.contains(k, na=False)
    if field in ("ledger", "both"):
        for k in kws:
            mask |= df["Ledger"].str.lower().str.strip("'").str.contains(k, na=False)
    if exclude:
        excl = pd.Series([False] * len(df), index=df.index)
        for ex in exclude:
            ex = ex.lower()
            excl |= df["Vch Type"].str.lower().str.strip("'").str.contains(ex, na=False)
            excl |= df["Ledger"].str.lower().str.strip("'").str.contains(ex, na=False)
        mask = mask & ~excl
    return df[mask]


def normalise_vch(v):
    return str(v).strip().strip("'").replace(" ", "").upper()


def calc_diff(co_cr, co_dr, pa_dr, pa_cr):
    try:
        co_cr = float(co_cr) if co_cr != "" else 0.0
    except:
        co_cr = 0.0
    try:
        co_dr = float(co_dr) if co_dr != "" else 0.0
    except:
        co_dr = 0.0
    try:
        pa_dr = float(pa_dr) if pa_dr != "" else 0.0
    except:
        pa_dr = 0.0
    try:
        pa_cr = float(pa_cr) if pa_cr != "" else 0.0
    except:
        pa_cr = 0.0
    if co_cr > 0 and pa_dr > 0:
        return round(co_cr - pa_dr, 2)
    if co_cr > 0:
        return round(co_cr - pa_cr, 2)
    if pa_dr > 0:
        return round(pa_dr, 2)
    if co_dr == 0 and pa_cr > 0:
        return round(pa_cr, 2)
    if co_dr > 0 or pa_cr > 0:
        return round(co_dr - pa_cr, 2)
    return 0.0


def build_merged_table(co_rows, party_rows):
    co = co_rows.copy().reset_index(drop=True)
    pa = party_rows.copy().reset_index(drop=True)
    co_used = [False] * len(co)
    pa_used = [False] * len(pa)
    rows = []
    pa_vch_map = {}
    for j, prow in pa.iterrows():
        nv = normalise_vch(prow.get("Vch No.", ""))
        if nv and nv not in ("", "NAN", "NONE"):
            pa_vch_map.setdefault(nv, []).append(j)
    # Pass 1: match by Vch No.
    for i, crow in co.iterrows():
        nv = normalise_vch(crow.get("Vch No.", ""))
        if not nv or nv in ("", "NAN", "NONE"):
            continue
        candidates = [j for j in pa_vch_map.get(nv, []) if not pa_used[j]]
        if not candidates:
            continue
        j = candidates[0]
        prow = pa.iloc[j]
        pa_used[j] = True
        co_used[i] = True
        co_dr = round(crow["Debit"], 2)
        co_cr = round(crow["Credit"], 2)
        pa_dr = round(prow["Debit"], 2)
        pa_cr = round(prow["Credit"], 2)
        rows.append(
            {
                "Date": crow["Date"],
                "Vch No.": crow["Vch No."],
                "Vch Type": crow["Vch Type"],
                "Ledger": crow["Ledger"],
                "Co. Debit": co_dr,
                "Co. Credit": co_cr,
                "Party Debit": pa_dr,
                "Party Credit": pa_cr,
                "Diff": calc_diff(co_cr, co_dr, pa_dr, pa_cr),
                "_src": "vch_match",
            }
        )
    # Pass 2: match by amount
    for i, crow in co.iterrows():
        if co_used[i]:
            continue
        co_cr = round(crow["Credit"], 2)
        co_dr = round(crow["Debit"], 2)
        matched_j = None
        if co_cr > 0:
            for j, prow in pa.iterrows():
                if pa_used[j]:
                    continue
                if round(prow["Debit"], 2) == co_cr:
                    matched_j = j
                    break
        if matched_j is None and co_dr > 0:
            for j, prow in pa.iterrows():
                if pa_used[j]:
                    continue
                if round(prow["Credit"], 2) == co_dr:
                    matched_j = j
                    break
        if matched_j is not None:
            prow = pa.iloc[matched_j]
            pa_used[matched_j] = True
            co_used[i] = True
            pa_dr = round(prow["Debit"], 2)
            pa_cr = round(prow["Credit"], 2)
            rows.append(
                {
                    "Date": crow["Date"],
                    "Vch No.": crow["Vch No."],
                    "Vch Type": crow["Vch Type"],
                    "Ledger": crow["Ledger"],
                    "Co. Debit": co_dr,
                    "Co. Credit": co_cr,
                    "Party Debit": pa_dr,
                    "Party Credit": pa_cr,
                    "Diff": calc_diff(co_cr, co_dr, pa_dr, pa_cr),
                    "_src": "amt_match",
                }
            )
        else:
            co_used[i] = True
            rows.append(
                {
                    "Date": crow["Date"],
                    "Vch No.": crow["Vch No."],
                    "Vch Type": crow["Vch Type"],
                    "Ledger": crow["Ledger"],
                    "Co. Debit": co_dr,
                    "Co. Credit": co_cr,
                    "Party Debit": "",
                    "Party Credit": "",
                    "Diff": co_cr if co_cr else co_dr,
                    "_src": "co_only",
                }
            )
    # Pass 3: remaining party rows
    for j, prow in pa.iterrows():
        if pa_used[j]:
            continue
        pa_dr = round(prow["Debit"], 2)
        pa_cr = round(prow["Credit"], 2)
        rows.append(
            {
                "Date": prow.get("Date", ""),
                "Vch No.": prow.get("Vch No.", ""),
                "Vch Type": prow["Vch Type"],
                "Ledger": prow["Ledger"],
                "Co. Debit": "",
                "Co. Credit": "",
                "Party Debit": pa_dr,
                "Party Credit": pa_cr,
                "Diff": pa_cr if pa_cr else -pa_dr,
                "_src": "party_only",
            }
        )
    if not rows:
        return pd.DataFrame(
            columns=[
                "Date",
                "Vch No.",
                "Vch Type",
                "Ledger",
                "Co. Debit",
                "Co. Credit",
                "Party Debit",
                "Party Credit",
                "Diff",
            ]
        )
    df = pd.DataFrame(rows).drop(columns=["_src"])

    def fmt_cell(x):
        if x == "" or x is None:
            return ""
        try:
            v = round(float(x), 2)
            return f"{v:,.2f}" if v != 0 else ""
        except:
            return str(x)

    for col in ["Co. Debit", "Co. Credit", "Party Debit", "Party Credit", "Diff"]:
        df[col] = df[col].apply(fmt_cell)
    return df


def fmt(val):
    try:
        v = round(float(val), 2)
        return "" if v == 0 else f"{v:,.2f}"
    except:
        return ""


def diff_fmt(val):
    try:
        v = round(float(val), 2)
        if v > 0:
            return f"+{v:,.2f}"
        elif v < 0:
            return f"{v:,.2f}"
        return "0.00"
    except:
        return "0.00"


# ──────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────
if file1 and file2:
    st.divider()
    df_raw1, fmt1 = read_file(file1)
    df_raw2, fmt2 = read_file(file2)
    if df_raw1 is None:
        st.error("Could not read Company Books.")
        st.stop()
    if df_raw2 is None:
        st.error("Could not read Party Statement.")
        st.stop()

    co_df, co_cols = extract_transactions(df_raw1, "Company Books", fmt1)
    party_df, party_cols = extract_transactions(df_raw2, "Party Statement", fmt2)
    if co_df is None or party_df is None:
        st.stop()

    co_close_amt, co_close_side = extract_closing_balance(df_raw1, fmt1, co_cols, file1)
    party_close_amt, party_close_side = extract_closing_balance(
        df_raw2, fmt2, party_cols, file2
    )

    c1, c2 = st.columns(2)
    c1.info(f"📘 **{file1.name}**")
    c2.info(f"📗 **{file2.name}**")

    with st.expander("🔍 Debug — Extracted Transactions"):
        t1, t2 = st.columns(2)
        with t1:
            st.markdown("**Company Books**")
            st.dataframe(co_df, use_container_width=True, hide_index=True)
            st.caption(f"Closing: {co_close_amt:,.2f} ({co_close_side})")
        with t2:
            st.markdown("**Party Statement**")
            st.dataframe(party_df, use_container_width=True, hide_index=True)
            st.caption(f"Closing: {party_close_amt:,.2f} ({party_close_side})")

    st.divider()

    recon_rows = []
    total_diff = 0.0

    purchase_grp = next((g for g in RECON_GROUPS if g["label"] == "Purchase"), None)
    if purchase_grp:
        _pu_party = search_df(
            party_df,
            purchase_grp["party_search"],
            purchase_grp["search_field"],
            purchase_grp.get("exclude", []),
        )
        purchase_party_cr = (
            round(_pu_party["Credit"].sum(), 2) if not _pu_party.empty else 0.0
        )
    else:
        purchase_party_cr = 0.0

    for grp in RECON_GROUPS:
        exclude = grp.get("exclude", [])
        co_rows = search_df(co_df, grp["co_search"], grp["search_field"], exclude)
        party_rows = search_df(
            party_df, grp["party_search"], grp["search_field"], exclude
        )
        co_dr = round(co_rows["Debit"].sum(), 2) if not co_rows.empty else 0.0
        co_cr = round(co_rows["Credit"].sum(), 2) if not co_rows.empty else 0.0
        party_dr = round(party_rows["Debit"].sum(), 2) if not party_rows.empty else 0.0
        party_cr = round(party_rows["Credit"].sum(), 2) if not party_rows.empty else 0.0

        if grp["label"] == "Opening Balance":
            cr_vs_dr = round((co_cr - party_dr) + (co_dr - party_cr), 2)
            diff = round((co_cr - party_dr) + (co_dr - party_cr), 2)
        elif grp["label"] == "Purchase":
            cr_vs_dr = round(co_cr - party_dr, 2)
            diff = round(cr_vs_dr + party_cr, 2)
        elif grp["label"] in ("Debit Note-QCR / GRN", "TDS"):
            cr_vs_dr = 0.0
            diff = round(-co_dr + party_cr, 2)
        elif grp["label"] == "Payment / Receipt / Bank / Journal":
            cr_vs_dr = round(co_cr, 2)
            diff = round(-(co_dr - co_cr - party_cr), 2)
        else:
            cr_vs_dr = round(co_cr - party_dr, 2)
            diff = round((co_cr + party_cr) - (party_dr + co_dr), 2)

        total_diff = round(total_diff + diff, 2)
        add_less = "Add" if diff > 0 else "Less" if diff < 0 else ""
        recon_rows.append(
            {
                "group": grp["label"],
                "add_less": add_less,
                "Co. Debit": co_dr,
                "Co. Credit": co_cr,
                "Party Debit": party_dr,
                "Party Credit": party_cr,
                "Cr vs Dr": cr_vs_dr,
                "Difference": diff,
                "co_rows": co_rows,
                "party_rows": party_rows,
            }
        )

    co_debit_close = co_close_amt if co_close_side == "debit" else 0.0
    co_credit_close = co_close_amt if co_close_side == "credit" else 0.0
    party_debit_close = party_close_amt if party_close_side == "debit" else 0.0
    party_credit_close = party_close_amt if party_close_side == "credit" else 0.0
    closing_diff = round(
        (co_debit_close + party_debit_close) - (co_credit_close + party_credit_close), 2
    )

    display_data = []
    for r in recon_rows:
        display_data.append(
            {
                "row_type": "recon",
                "Add/Less": r["add_less"],
                "Description": f"Diff in {r['group']}",
                "Co. Debit": fmt(r["Co. Debit"]),
                "Co. Credit": fmt(r["Co. Credit"]),
                "Party Debit": fmt(r["Party Debit"]),
                "Party Credit": fmt(r["Party Credit"]),
                "Co.Cr vs Party.Dr": (
                    diff_fmt(r["Cr vs Dr"]) if r["Cr vs Dr"] != 0 else "0"
                ),
                "Difference": diff_fmt(r["Difference"]),
            }
        )

    total_co_dr = round(sum(r["Co. Debit"] for r in recon_rows), 2)
    total_co_cr = round(sum(r["Co. Credit"] for r in recon_rows), 2)
    total_party_dr = round(sum(r["Party Debit"] for r in recon_rows), 2)
    total_party_cr = round(sum(r["Party Credit"] for r in recon_rows), 2)
    display_data.append(
        {
            "row_type": "total",
            "Add/Less": "",
            "Description": "Total Diff",
            "Co. Debit": fmt(total_co_dr),
            "Co. Credit": fmt(total_co_cr),
            "Party Debit": fmt(total_party_dr),
            "Party Credit": fmt(total_party_cr),
            "Co.Cr vs Party.Dr": "",
            "Difference": diff_fmt(total_diff),
        }
    )
    display_data.append(
        {
            "row_type": "closing",
            "Add/Less": "",
            "Description": "Diff in closing balance",
            "Co. Debit": fmt(co_close_amt) if co_close_side == "debit" else "",
            "Co. Credit": fmt(co_close_amt) if co_close_side == "credit" else "",
            "Party Debit": fmt(party_close_amt) if party_close_side == "debit" else "",
            "Party Credit": (
                fmt(party_close_amt) if party_close_side == "credit" else ""
            ),
            "Co.Cr vs Party.Dr": "",
            "Difference": diff_fmt(closing_diff),
        }
    )

    display_df = pd.DataFrame(display_data)

    def style_recon(row):
        rt = display_df.loc[row.name, "row_type"]
        if rt == "total":
            return [
                "background-color:#dbeafe;font-weight:bold;border:2px solid #3b82f6;"
            ] * len(row)
        elif rt == "closing":
            return [
                "background-color:#cce5ff;font-weight:bold;border:2px solid #4a9eda"
            ] * len(row)
        else:
            try:
                v = float(
                    str(display_df.loc[row.name, "Difference"])
                    .replace("+", "")
                    .replace(",", "")
                )
                color = "#ffe0e0" if v < 0 else ("#e8f5e9" if v > 0 else "#ffffff")
                return [f"background-color:{color};border:1px solid #ccc"] * len(row)
            except:
                return ["border:1px solid #eee"] * len(row)

    display_cols = [
        "Add/Less",
        "Description",
        "Co. Debit",
        "Co. Credit",
        "Party Debit",
        "Party Credit",
        "Co.Cr vs Party.Dr",
        "Difference",
    ]

    st.subheader("📊 Reconciliation Statement")
    st.caption(
        "💡 Click any row to drill down — line-by-line comparison in company book style"
    )

    # CSV Export
    recon_export_rows = []
    for r in recon_rows:
        recon_export_rows.append(
            {
                "Add/Less": r["add_less"],
                "Description": f"Diff in {r['group']}",
                "Co. Debit": fmt(r["Co. Debit"]),
                "Co. Credit": fmt(r["Co. Credit"]),
                "Party Debit": fmt(r["Party Debit"]),
                "Party Credit": fmt(r["Party Credit"]),
                "Co.Cr vs Party.Dr": (
                    diff_fmt(r["Cr vs Dr"]) if r["Cr vs Dr"] != 0 else "0"
                ),
                "Difference": diff_fmt(r["Difference"]),
            }
        )
    recon_export_rows.append(
        {
            "Add/Less": "",
            "Description": "Total Diff",
            "Co. Debit": fmt(total_co_dr),
            "Co. Credit": fmt(total_co_cr),
            "Party Debit": fmt(total_party_dr),
            "Party Credit": fmt(total_party_cr),
            "Co.Cr vs Party.Dr": "",
            "Difference": diff_fmt(total_diff),
        }
    )
    recon_export_rows.append(
        {
            "Add/Less": "",
            "Description": "Diff in closing balance",
            "Co. Debit": fmt(co_close_amt) if co_close_side == "debit" else "",
            "Co. Credit": fmt(co_close_amt) if co_close_side == "credit" else "",
            "Party Debit": fmt(party_close_amt) if party_close_side == "debit" else "",
            "Party Credit": (
                fmt(party_close_amt) if party_close_side == "credit" else ""
            ),
            "Co.Cr vs Party.Dr": "",
            "Difference": diff_fmt(closing_diff),
        }
    )
    csv_out = pd.DataFrame(recon_export_rows, columns=display_cols).to_csv(index=False)
    st.download_button(
        label="⬇️ Download Reconciliation CSV",
        data=csv_out,
        file_name="reconciliation_statement.csv",
        mime="text/csv",
    )

    event = st.dataframe(
        display_df[display_cols].style.apply(style_recon, axis=1),
        use_container_width=True,
        hide_index=True,
        on_select="rerun",
        selection_mode="single-row",
    )

    selected = event.selection.rows
    if selected:
        sel_idx = selected[0]
        sel_rt = display_df.iloc[sel_idx]["row_type"]
        sel_desc = display_df.iloc[sel_idx]["Description"]
        st.divider()
        st.subheader(f"🔍 {sel_desc}")

        if sel_rt == "recon" and sel_idx < len(recon_rows):
            grp = recon_rows[sel_idx]
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Co. Debit", fmt(grp["Co. Debit"]))
            m2.metric("Co. Credit", fmt(grp["Co. Credit"]))
            m3.metric("Party Debit", fmt(grp["Party Debit"]))
            m4.metric("Party Credit", fmt(grp["Party Credit"]))
            m5.metric("Difference", diff_fmt(grp["Difference"]))
            st.divider()
            merged = build_merged_table(grp["co_rows"], grp["party_rows"])
            st.markdown("#### 📋 Line-by-line Comparison (Company Book Style)")
            st.caption("🟢 Matched  |  🔴 Difference  |  🟡 Not present in one side")
            if not merged.empty:

                def color_merged_row(row):
                    diff_str = str(row.get("Diff", ""))
                    if row["Co. Credit"] == "" or row["Party Debit"] == "":
                        return ["background-color:#fff3cd"] * len(row)
                    try:
                        v = float(diff_str.replace(",", "").replace("+", ""))
                        return [
                            (
                                "background-color:#e8f5e9"
                                if v == 0
                                else "background-color:#ffe0e0"
                            )
                        ] * len(row)
                    except:
                        return ["background-color:#fff3cd"] * len(row)

                st.dataframe(
                    merged.style.apply(color_merged_row, axis=1),
                    use_container_width=True,
                    hide_index=True,
                )

                def safe_sum(col):
                    total = 0.0
                    for v in merged[col]:
                        try:
                            total += float(str(v).replace(",", ""))
                        except:
                            pass
                    return total

                t1, t2, t3, t4, t5 = st.columns(5)
                t1.metric("Co. Debit", f"{safe_sum('Co. Debit'):,.2f}")
                t2.metric("Co. Credit", f"{safe_sum('Co. Credit'):,.2f}")
                t3.metric("Party Debit", f"{safe_sum('Party Debit'):,.2f}")
                t4.metric("Party Credit", f"{safe_sum('Party Credit'):,.2f}")
                t5.metric("Net Diff", f"{safe_sum('Diff'):,.2f}")
            else:
                st.success("✅ No transactions found for this group.")

        elif sel_rt == "total":
            st.info("Sum of all reconciliation differences.")
            st.dataframe(
                pd.DataFrame(
                    [
                        {
                            "Group": g["group"],
                            "Add/Less": g["add_less"],
                            "Co.Cr vs Party.Dr": diff_fmt(g["Cr vs Dr"]),
                            "Difference": diff_fmt(g["Difference"]),
                        }
                        for g in recon_rows
                    ]
                ),
                use_container_width=True,
                hide_index=True,
            )

        elif sel_rt == "closing":
            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown("**🏢 Company Books — Closing Balance**")
                st.metric(
                    f"Closing Balance ({'Debit' if co_close_side=='debit' else 'Credit'})",
                    f"{co_close_amt:,.2f}",
                )
                st.dataframe(
                    co_df[["Date", "Vch No.", "Vch Type", "Ledger", "Debit", "Credit"]],
                    use_container_width=True,
                    hide_index=True,
                )
            with col_b:
                st.markdown("**🤝 Party Statement — Closing Balance**")
                st.metric(
                    f"Closing Balance ({'Debit' if party_close_side=='debit' else 'Credit'})",
                    f"{party_close_amt:,.2f}",
                )
                available = [
                    c
                    for c in [
                        "Date",
                        "Vch No.",
                        "Vch Type",
                        "Ledger",
                        "Debit",
                        "Credit",
                    ]
                    if c in party_df.columns
                ]
                st.dataframe(
                    party_df[available], use_container_width=True, hide_index=True
                )

    st.divider()
    st.subheader("📊 Summary")
    g1, g2, g3 = st.columns(3)
    g1.metric(f"Co. Closing ({co_close_side.title()})", f"{co_close_amt:,.2f}")
    g2.metric(f"Party Closing ({party_close_side.title()})", f"{party_close_amt:,.2f}")
    g3.metric("Closing Diff", diff_fmt(closing_diff))

elif file1 and not file2:
    st.warning("⬆️ Please upload the Party Statement file too.")
elif file2 and not file1:
    st.warning("⬆️ Please upload the Company Books file too.")
else:
    st.info("⬆️ Please upload both files to start comparison.")
