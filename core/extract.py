#!/usr/bin/env python3
"""context_extract engine CLI.

Read a raw source file (financial statement workbook, business ledger CSV/XLSX,
contract DOCX/TXT/MD) and compress it into high-density structured Markdown +
JSON for downstream M1-M4 modules. Deterministic extraction only: structure,
entities, key line items, and light cross-checks (勾稽校验). Semantic
"compress each clause to one sentence" is the calling agent's job, guided by the
system-prompt section this plugin injects.

Usage:
  python core/extract.py --input model.xlsx \
      --out-json report.json --out-md report.md [--targets 货币资金,净利润] \
      [--max-rows 200]

Prints one JSON object to stdout (the same payload written to --out-json).
Progress and errors go to stderr.

Supported inputs: .xlsx/.xlsm (openpyxl), .csv, .docx (python-docx), .txt/.md.
PDF is NOT supported here (the bundled pdfplumber is broken in this runtime and
pymupdf is absent); convert PDF first via LibreOffice Kit.
"""
from __future__ import annotations

import argparse
import csv
import json
import math
import os
import re
import sys
from datetime import date, datetime

from openpyxl import load_workbook

ENGINE_VERSION = "0.1.0"

CHARS_PER_TOKEN = 4

# --------------------------------------------------------------------------- #
# helpers (adapted from dsh-office-audit core/auditor/util.py)
# --------------------------------------------------------------------------- #

CJK_RE = re.compile(r"[\u4e00-\u9fff]")
YEAR_RE = re.compile(r"(?<!\d)(20\d{2})(?!\d)")


def normalize_label(text) -> str:
    if not isinstance(text, str):
        return ""
    text = text.replace("\u3000", " ").replace("\xa0", " ")
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"[（(].*?[）)]", "", text)
    text = text.rstrip(":：")
    return text


def is_number(value) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def to_number(value):
    if is_number(value):
        try:
            f = float(value)
        except (TypeError, ValueError):
            return None
        if math.isnan(f) or math.isinf(f):
            return None
        return f
    if isinstance(value, str):
        text = value.strip().replace(",", "").replace("，", "")
        if text in ("", "-", "—", "–", "N/A", "n/a", "NA", "不适用", "无"):
            return None
        pct = text.endswith("%") or text.endswith("％")
        if pct:
            text = text[:-1]
        try:
            f = float(text)
        except ValueError:
            return None
        return f / 100.0 if pct else f
    return None


def period_of_header(text):
    if text is None:
        return None
    s = str(text).strip()
    if s == "":
        return None
    ym = YEAR_RE.search(s)
    year = int(ym.group(1)) if ym else None
    if year is None:
        return None
    q = re.search(r"(?:Q|q)\s*([1-4])", s)
    if q:
        return ("quarter", year, int(q.group(1)))
    m = re.search(r"(?:第?\s*([1-4])\s*季)", s)
    if m:
        return ("quarter", year, int(m.group(1)))
    month = re.search(r"(\d{1,2})\s*月", s)
    if month and 1 <= int(month.group(1)) <= 12:
        return ("month", year, int(month.group(1)))
    return ("year", year)


def period_label(period) -> str:
    if period is None:
        return "值"
    if period[0] == "quarter" and len(period) > 2:
        return f"{period[1]}Q{period[2]}"
    if period[0] == "month" and len(period) > 2:
        return f"{period[1]}-{period[2]:02d}"
    return str(period[1])


def fmt_num(value) -> str:
    num = to_number(value)
    if num is None:
        return "" if value is None else str(value)
    if abs(num) >= 1e8:
        return f"{num / 1e8:,.2f}亿"
    if abs(num) >= 1e4:
        return f"{num / 1e4:,.2f}万"
    if abs(num) >= 1000:
        return f"{num:,.2f}"
    if num == int(num):
        return str(int(num))
    return f"{num:,.4g}"


# Accounting labels used for key-item extraction and cross-checks.
LABELS = {
    "assets_total": ("资产总计", "资产合计", "总资产", "资产总额"),
    "liabilities_total": ("负债合计", "负债总计", "负债总额"),
    "equity_total": ("所有者权益合计", "股东权益合计", "所有者权益总计", "股东权益总计", "权益合计"),
    "cash_bs": ("货币资金", "现金及现金等价物", "现金及银行存款", "银行存款"),
    "cash_cf_end": ("期末现金及现金等价物余额", "期末现金余额", "期末现金"),
    "cash_cf_begin": ("期初现金及现金等价物余额", "期初现金余额", "期初现金"),
    "cash_cf_net": ("现金及现金等价物净增加额", "现金净增加额", "现金及现金等价物净变动额"),
    "cf_operating": ("经营活动产生的现金流量净额", "经营活动现金流量净额", "经营活动净现金流"),
    "cf_investing": ("投资活动产生的现金流量净额", "投资活动现金流量净额", "投资活动净现金流"),
    "cf_financing": ("筹资活动产生的现金流量净额", "筹资活动现金流量净额", "筹资活动净现金流"),
    "net_income": ("归属于母公司股东的净利润", "归母净利润", "净利润", "税后利润"),
    "revenue": ("营业收入", "营业总收入", "销售收入", "营收"),
    "total_costs": ("营业总成本", "营业成本"),
    "retained_earnings": ("留存收益", "未分配利润", "留存利润"),
}

# Order in which key items are presented (most important first).
KEY_ITEM_ORDER = [
    "revenue", "total_costs", "net_income",
    "assets_total", "liabilities_total", "equity_total", "cash_bs",
    "cf_operating", "cf_investing", "cf_financing", "cash_cf_net",
    "cash_cf_begin", "cash_cf_end", "retained_earnings",
]

# Substring labels that must NOT appear when matching a key (avoids cross-lining
# e.g. BS 货币资金 vs CF 期初/期末现金及现金等价物余额).
KEY_EXCLUDE = {
    "cash_bs": ("期初", "期末", "余额", "净", "增加", "减少", "流量", "变动"),
    "net_income": ("少数股东",),
}

DERIVED_SHEET_WORDS = (
    "common-size", "commonsize", "common size", "结构", "占比", "垂直",
    "同比", "勾稽校验", "模型审计", "敏感性", "sensitivity", "审计",
)

FINANCIAL_MARKERS = (
    "资产总计", "资产合计", "总资产", "负债合计", "所有者权益", "股东权益",
    "货币资金", "净利润", "营业收入", "经营活动产生的现金流量", "现金流量净额",
)

CONTRACT_MARKERS = (
    "甲方", "乙方", "合同", "协议", "条款", "签署", "争议解决", "违约责任",
    "第", "条", "章",
)

COMPANY_RE = re.compile(
    r"[\u4e00-\u9fffA-Za-z0-9（）()]{2,}(?:有限责任公司|股份有限公司|有限合伙|合伙企业|集团|事务所|公司|银行|证券|基金|投资)"
)

DATE_RE = re.compile(r"(?:20\d{2}\s*[-年/.]\s*\d{1,2}\s*[-月/.]\s*\d{1,2})|(?:20\d{2}\s*年\s*\d{1,2}\s*月)")

CURRENCY_RE = re.compile(r"[¥$€£￥]|\d[\d,，.]*(?:元|万元|亿元|千元|百万|亿|万)")


def is_derived_sheet_name(name) -> bool:
    norm = str(name).lower().replace(" ", "")
    return any(word in norm for word in DERIVED_SHEET_WORDS)


def detect_unit(texts) -> str | None:
    joined = " ".join(str(t) for t in texts if t)
    if "亿元" in joined:
        return "人民币亿元"
    if "万元" in joined:
        return "人民币万元"
    if "千元" in joined:
        return "人民币千元"
    if "百万元" in joined or "百万" in joined:
        return "人民币百万元"
    if "元" in joined and not re.search(r"亿元|万元|千元|百万元", joined):
        return "人民币元"
    return None


# --------------------------------------------------------------------------- #
# spreadsheet model
# --------------------------------------------------------------------------- #

HEADER_SCAN_ROWS = 15


class Sheet:
    def __init__(self, ws):
        self.ws = ws
        self.name = ws.title
        self.max_row = ws.max_row or 0
        self.max_col = ws.max_column or 0
        self.header_row = self._detect_header_row()
        self.period_columns = self._detect_periods()
        self._merged = list(ws.merged_cells.ranges) if ws.merged_cells else []

    def cell(self, r, c):
        return self.ws.cell(row=r, column=c).value

    def _detect_header_row(self):
        best_row, best_score = None, 0
        limit = min(self.max_row, HEADER_SCAN_ROWS)
        for row in range(1, limit + 1):
            score = 0
            for col in range(1, self.max_col + 1):
                v = self.cell(row, col)
                if isinstance(v, str) and len(v.strip()) > 0 and not v.startswith("="):
                    # A header cell carries a label (CJK/latin) OR a period/date (digits).
                    if re.search(r"[A-Za-z\u4e00-\u9fff]", v) or re.search(r"\d", v):
                        score += 1
            if score > best_score:
                best_row, best_score = row, score
        return best_row if best_score >= 2 else None

    def _detect_periods(self):
        periods = {}
        if self.header_row is None:
            return periods
        for col in range(1, self.max_col + 1):
            p = period_of_header(self.cell(self.header_row, col))
            if p is not None:
                periods[col] = p
        return periods

    def _merged_lookup(self, row, col):
        for rng in self._merged:
            if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
                return self.cell(rng.min_row, rng.min_col)
        return None

    def label(self, row) -> str:
        for col in range(1, min(self.max_col, 3) + 1):
            v = self.cell(row, col)
            if isinstance(v, str) and v.strip() and not v.startswith("="):
                return v.strip()
            if v is None:
                anchor = self._merged_lookup(row, col)
                if anchor is not None and isinstance(anchor, str) and anchor.strip():
                    return anchor.strip()
        return ""

    def label_rows(self):
        out = []
        for row in range(1, self.max_row + 1):
            label = self.label(row)
            if label:
                out.append((row, label))
        return out

    def values_for(self, row):
        vals = {}
        for col, period in self.period_columns.items():
            vals[period_label(period)] = self.cell(row, col)
        # Also capture non-period numeric columns (e.g. a single "金额" column).
        return vals


# --------------------------------------------------------------------------- #
# extraction
# --------------------------------------------------------------------------- #

def extract_excel(path):
    wb = load_workbook(path, data_only=True, read_only=False)
    sheets = []
    key_items = []
    tables = []
    all_texts = []
    formula_note = False

    for ws in wb.worksheets:
        if is_derived_sheet_name(ws.title):
            continue
        sheet = Sheet(ws)
        if sheet.header_row is None:
            continue
        # collect text for unit/entity detection
        for r in range(1, min(sheet.max_row, HEADER_SCAN_ROWS) + 1):
            for c in range(1, sheet.max_col + 1):
                v = sheet.cell(r, c)
                if isinstance(v, str):
                    all_texts.append(v)
        # key items from labels
        for key in KEY_ITEM_ORDER:
            words = LABELS[key]
            excludes = KEY_EXCLUDE.get(key, ())
            for row, label in sheet.label_rows():
                norm = normalize_label(label)
                if any(bad in norm for bad in excludes):
                    continue
                if any(w in norm for w in words):
                    values = {}
                    for col, period in sheet.period_columns.items():
                        values[period_label(period)] = sheet.cell(row, col)
                    # dedupe by (key, sheet) — keep first match per sheet
                    if not any(it.get("key") == key and it.get("sheet") == sheet.name for it in key_items):
                        key_items.append({
                            "key": key,
                            "label": label,
                            "sheet": sheet.name,
                            "values": values,
                        })
                    break
        # ledger / generic table dump (non-financial sheets)
        if len(sheet.period_columns) >= 1 or _looks_like_table(sheet):
            table = _dump_table(sheet)
            if table is not None:
                tables.append(table)

    for row in range(1, min(wb.active.max_row or 0, 50) + 1):
        for col in range(1, min(wb.active.max_column or 0, 50) + 1):
            v = wb.active.cell(row=row, column=col).value
            if isinstance(v, str) and v.startswith("="):
                formula_note = True
                break

    unit = detect_unit(all_texts)
    return {
        "sheets": sheets,
        "key_items": key_items,
        "tables": tables,
        "unit": unit,
        "formula_note": "公式单元格无缓存值时按空处理；如需精确数值请先用 LibreOffice 重算。" if formula_note else None,
    }


def _looks_like_table(sheet) -> bool:
    if sheet.header_row is None:
        return False
    start = sheet.header_row + 1
    numeric = 0
    for row in range(start, min(sheet.max_row, start + 30) + 1):
        for col in range(1, sheet.max_col + 1):
            if is_number(sheet.cell(row, col)):
                numeric += 1
    return numeric >= 3


def _dump_table(sheet, max_rows=200):
    if sheet.header_row is None:
        return None
    headers = []
    for col in range(1, sheet.max_col + 1):
        h = sheet.cell(sheet.header_row, col)
        if h is None:
            continue
        headers.append(str(h).strip())
    if not headers:
        return None
    rows = []
    for row in range(sheet.header_row + 1, min(sheet.max_row, sheet.header_row + 1 + max_rows) + 1):
        vals = []
        has_value = False
        for col in range(1, len(headers) + 1):
            v = sheet.cell(row, col)
            if v is not None and str(v).strip() != "":
                has_value = True
            vals.append(v)
        if has_value:
            rows.append(vals)
    return {"sheet": sheet.name, "headers": headers, "rows": rows, "truncated": sheet.max_row > sheet.header_row + 1 + max_rows}


def extract_csv(path):
    rows = []
    with open(path, "r", encoding="utf-8-sig", errors="replace", newline="") as f:
        reader = csv.reader(f)
        for row in reader:
            rows.append(row)
    if not rows:
        return {"key_items": [], "tables": [], "unit": None, "all_texts": []}
    headers = [str(h).strip() for h in rows[0]]
    data = []
    for row in rows[1:201]:
        data.append(row)
    return {
        "key_items": [],
        "tables": [{"sheet": os.path.basename(path), "headers": headers, "rows": data, "truncated": len(rows) > 201}],
        "unit": detect_unit([" ".join(headers)]),
        "all_texts": [" ".join(headers)],
    }


def extract_docx(path):
    try:
        import docx  # python-docx
    except ImportError:
        return {"paragraphs": [], "tables": [], "error": "python-docx 不可用"}
    doc = docx.Document(path)
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    tables = []
    for t in doc.tables:
        headers = [c.text.strip() for c in t.rows[0].cells] if t.rows else []
        rows = []
        for r in t.rows[1:201]:
            rows.append([c.text.strip() for c in r.cells])
        tables.append({"headers": headers, "rows": rows})
    return {"paragraphs": paragraphs, "tables": tables, "unit": detect_unit(paragraphs)}


def extract_text(path):
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        text = f.read()
    return {"paragraphs": [text], "tables": [], "unit": detect_unit([text])}


# --------------------------------------------------------------------------- #
# type detection
# --------------------------------------------------------------------------- #

def classify(ext, key_items, paragraphs, all_texts):
    joined = " ".join(paragraphs[:200]) if paragraphs else ""
    has_financial = any(w in joined for w in FINANCIAL_MARKERS) or len(key_items) >= 2
    has_contract = any(w in joined for w in CONTRACT_MARKERS)
    if ext in (".xlsx", ".xlsm", ".csv"):
        if has_financial:
            return "financial_statement", "财务报表"
        return "ledger", "业务台账"
    if has_contract:
        return "contract", "法律合同"
    return "text", "文本"


# --------------------------------------------------------------------------- #
# clauses (contract)
# --------------------------------------------------------------------------- #

CLAUSE_MARKER_RE = re.compile(
    r"第\s*[一二三四五六七八九十百零\d]+\s*[条章]"
    r"|Article\s+\d+"
    r"|Clause\s+\d+"
    r"|SECTION\s+\d+"
    r"|^\s*\d+\s*[、.．)）]",
    re.MULTILINE,
)


def split_clauses(text, max_len=120):
    """Split a contract text into numbered clauses (deterministic, structural)."""
    matches = list(CLAUSE_MARKER_RE.finditer(text))
    if not matches:
        return [{"no": "全文", "text": text.strip()[:max_len]}]
    clauses = []
    pre = text[: matches[0].start()].strip()
    if pre:
        clauses.append({"no": "前言", "text": pre[:max_len]})
    for i, m in enumerate(matches):
        marker = m.group(0).strip()
        body_start = m.end()
        body_end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        body = text[body_start:body_end].strip()
        clauses.append({"no": marker, "text": body[:max_len]})
    return clauses


# --------------------------------------------------------------------------- #
# entities
# --------------------------------------------------------------------------- #

def extract_entities(paragraphs, tables):
    text_pool = "\n".join(paragraphs) if paragraphs else ""
    for t in tables:
        text_pool += "\n" + " ".join(t.get("headers", []))
        for row in t.get("rows", []):
            text_pool += "\n" + " ".join(str(c) for c in row if c is not None)
    companies = sorted(set(COMPANY_RE.findall(text_pool)))[:20]
    dates = sorted(set(DATE_RE.findall(text_pool)))[:20]
    amounts = sorted(set(CURRENCY_RE.findall(text_pool)))[:20]
    return {"companies": companies, "dates": dates, "amounts": amounts}


# --------------------------------------------------------------------------- #
# cross-check (勾稽校验)
# --------------------------------------------------------------------------- #

def cross_check(key_items):
    alerts = []
    by_key = {}
    for it in key_items:
        by_key.setdefault(it["key"], []).append(it)

    def num(it, label):
        v = it["values"].get(label)
        return to_number(v)

    # 1. balance-sheet: 资产总计 ≈ 负债合计 + 所有者权益合计
    for assets in by_key.get("assets_total", []):
        for liab in by_key.get("liabilities_total", []):
            for eq in by_key.get("equity_total", []):
                for label in set(list(assets["values"]) + list(liab["values"]) + list(eq["values"])):
                    a, l, e = num(assets, label), num(liab, label), num(eq, label)
                    if a is None or l is None or e is None:
                        continue
                    diff = a - (l + e)
                    if abs(diff) > max(0.001 * max(abs(a), 1.0), 1.0):
                        alerts.append({
                            "tag": "alert",
                            "check": "balance-sheet",
                            "period": label,
                            "message": f"资产总计 ≠ 负债合计 + 所有者权益合计（差 {fmt_num(diff)}）",
                            "detail": f"资产 {fmt_num(a)} vs 负债 {fmt_num(l)} + 权益 {fmt_num(e)}",
                        })

    # 2. cash tie-out: BS 货币资金 vs CF 期末现金及现金等价物
    for cash in by_key.get("cash_bs", []):
        for cf_end in by_key.get("cash_cf_end", []):
            for label in set(list(cash["values"]) + list(cf_end["values"])):
                a, b = num(cash, label), num(cf_end, label)
                if a is None or b is None:
                    continue
                if abs(a - b) > max(0.001 * max(abs(a), 1.0), 1.0):
                    alerts.append({
                        "tag": "alert",
                        "check": "cash-tieout",
                        "period": label,
                        "message": f"货币资金 ≠ 期末现金及现金等价物（差 {fmt_num(a - b)}）",
                        "detail": f"BS {fmt_num(a)} vs CF {fmt_num(b)}",
                    })

    return alerts


# --------------------------------------------------------------------------- #
# markdown builder
# --------------------------------------------------------------------------- #

def build_markdown(payload):
    lines = []
    lines.append(f"# {payload['type_label']} · {os.path.basename(payload['input'])}")
    meta = [f"类型: {payload['type']}"]
    if payload.get("unit"):
        meta.append(f"单位: {payload['unit']}")
    lines.append(" · ".join(meta))

    # key items table
    if payload.get("key_items"):
        # collect period columns
        cols = []
        for it in payload["key_items"]:
            for label in it["values"]:
                if label not in cols:
                    cols.append(label)
        header = "| 指标 |" + "|".join(cols) + "|"
        sep = "|---|" + "|".join(["---"] * len(cols)) + "|"
        lines.append("")
        lines.append("## 关键指标")
        lines.append(header)
        lines.append(sep)
        for it in payload["key_items"]:
            row = [it["label"]]
            for col in cols:
                row.append(fmt_num(it["values"].get(col)))
            lines.append("| " + " | ".join(row) + " |")

    # tables (ledger)
    if payload.get("tables"):
        lines.append("")
        lines.append("## 数据表")
        for t in payload["tables"]:
            lines.append(f"### {t['sheet']}")
            lines.append("| " + " | ".join(t["headers"]) + " |")
            lines.append("|" + "|".join(["---"] * len(t["headers"])) + "|")
            for row in t["rows"]:
                cells = [fmt_num(c) if is_number(c) else str(c) for c in row]
                # pad to header length
                while len(cells) < len(t["headers"]):
                    cells.append("")
                lines.append("| " + " | ".join(cells[: len(t["headers"])]) + " |")
            if t.get("truncated"):
                lines.append(f"… 截断（仅前 200 行）")

    # clauses (contract)
    if payload.get("clauses"):
        lines.append("")
        lines.append("## 条款")
        lines.append("| 条款 | 内容 |")
        lines.append("|---|---|")
        for c in payload["clauses"]:
            lines.append(f"| {c['no']} | {c['text']} |")

    # entities
    ent = payload.get("entities", {})
    if ent.get("companies") or ent.get("dates") or ent.get("amounts"):
        lines.append("")
        lines.append("## 实体")
        if ent.get("companies"):
            lines.append(f"- 主体: {'；'.join(ent['companies'])}")
        if ent.get("dates"):
            lines.append(f"- 日期: {'；'.join(ent['dates'])}")
        if ent.get("amounts"):
            lines.append(f"- 金额: {'；'.join(ent['amounts'])}")

    # alerts
    if payload.get("alerts"):
        lines.append("")
        lines.append("## 勾稽提示")
        for a in payload["alerts"]:
            lines.append(f"- <alert> [{a['check']}] {a['period']}: {a['message']} — {a['detail']}")

    if payload.get("note"):
        lines.append("")
        lines.append(f"> {payload['note']}")

    return "\n".join(lines)


def estimate_tokens(text) -> int:
    return math.ceil(len(text) / CHARS_PER_TOKEN)


# --------------------------------------------------------------------------- #
# main
# --------------------------------------------------------------------------- #

def run(path, targets=None, max_rows=200):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        result = extract_excel(path)
        paragraphs = []
        key_items = result["key_items"]
        tables = result["tables"]
        all_texts = result.get("unit", "") or ""
    elif ext == ".csv":
        result = extract_csv(path)
        paragraphs = result.get("all_texts", [])
        key_items = result["key_items"]
        tables = result["tables"]
        all_texts = result.get("unit", "") or ""
    elif ext == ".docx":
        result = extract_docx(path)
        paragraphs = result.get("paragraphs", [])
        key_items = []
        tables = result.get("tables", [])
        all_texts = result.get("unit", "") or ""
    elif ext in (".txt", ".md"):
        result = extract_text(path)
        paragraphs = result.get("paragraphs", [])
        key_items = []
        tables = result.get("tables", [])
        all_texts = result.get("unit", "") or ""
    else:
        return {
            "ok": False,
            "error": f"unsupported file type '{ext}'; supported: .xlsx/.xlsm/.csv/.docx/.txt/.md "
                      "(PDF 需先用 LibreOffice Kit 转成文本或 xlsx)",
        }

    if result.get("error"):
        return {"ok": False, "error": result["error"]}

    file_type, type_label = classify(ext, key_items, paragraphs, [all_texts])

    # targets filter: keep only key items whose label matches a target word
    if targets:
        filtered = []
        for it in key_items:
            norm = normalize_label(it["label"])
            if any(t in norm for t in targets):
                filtered.append(it)
        key_items = filtered

    clauses = []
    if file_type == "contract":
        full_text = "\n".join(paragraphs)
        clauses = split_clauses(full_text)
        if not clauses:
            clauses = [{"no": "全文", "text": full_text[:200]}]

    entities = extract_entities(paragraphs, tables)
    alerts = cross_check(key_items) if file_type == "financial_statement" else []

    payload = {
        "ok": True,
        "engine_version": ENGINE_VERSION,
        "input": path,
        "type": file_type,
        "type_label": type_label,
        "unit": result.get("unit"),
        "entities": entities,
        "key_items": key_items,
        "tables": tables,
        "clauses": clauses,
        "alerts": alerts,
        "note": result.get("formula_note"),
    }
    payload["markdown"] = build_markdown(payload)
    payload["token_estimate"] = estimate_tokens(payload["markdown"])
    return payload


def main():
    parser = argparse.ArgumentParser(description="context_extract engine")
    parser.add_argument("--input", required=True)
    parser.add_argument("--out-json", required=True)
    parser.add_argument("--out-md", required=True)
    parser.add_argument("--targets", default="")
    parser.add_argument("--max-rows", type=int, default=200)
    args = parser.parse_args()

    targets = [t.strip() for t in args.targets.split(",") if t.strip()] if args.targets else None
    payload = run(args.input, targets=targets, max_rows=args.max_rows)

    # persist report files
    os.makedirs(os.path.dirname(os.path.abspath(args.out_json)) or ".", exist_ok=True)
    with open(args.out_json, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    with open(args.out_md, "w", encoding="utf-8") as f:
        f.write(payload.get("markdown", "") if payload.get("ok") else f"# 萃取失败\n\n{payload.get('error', '')}")

    print(json.dumps(payload, ensure_ascii=False))


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        print(json.dumps({"ok": False, "error": f"{type(exc).__name__}: {exc}"}, ensure_ascii=False))
        sys.exit(0)
