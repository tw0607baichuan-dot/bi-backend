#!/usr/bin/env python3
"""
Phase 2.2 — Excel 解析模组

对外两个函式:
  parse_filename(filename, default_year=None) -> {product_line, week_start, week_end}
  parse_excel(filepath)                        -> list[dict]  (每个 dict = 一笔退费,7 栏)

例外阶层(各自 raise 不同 exception,方便上层回不同 HTTP code):
  ExcelParseError        — 基底
    FilenameParseError   — 档名格式不符
    FileCorruptError     — 档案损坏 / 无法开启
    MissingColumnError   — 缺必要栏位
"""
import datetime
import os
import re

import pandas as pd


# ---------- 例外 ----------
class ExcelParseError(Exception):
    pass


class FilenameParseError(ExcelParseError):
    pass


class FileCorruptError(ExcelParseError):
    pass


class MissingColumnError(ExcelParseError):
    pass


# ---------- 栏位对照 ----------
# DB 栏 -> 可接受的 Excel 表头别名(按优先序)
COLUMN_ALIASES = {
    "date_iso":          ["日期", "退费日期", "date", "Date"],
    "merchant_tg":       ["商户", "商户TG", "商户名", "merchant"],
    "merchant_order_no": ["商户单号", "商户订单号", "merchant_order_no"],
    "platform_order_no": ["平台单号", "平台订单号", "platform_order_no"],
    "amount":            ["投诉金额", "金额", "退费金额", "amount"],
    "payment_type":      ["支付类型", "支付方式", "payment_type"],
    "platform_id":       ["平台ID", "平台id", "平台", "platform_id"],
}

# Excel 序列号基准日(Excel 的 1900 闰年 bug → 用 1899-12-30 当 day 0)
_EXCEL_EPOCH = datetime.datetime(1899, 12, 30)


# ---------- 档名解析 ----------
def parse_filename(filename, default_year=None):
    """
    "5_17_5_23悦达.xlsx" -> {product_line, week_start, week_end}
    年份从 default_year 带入(由内容第一笔日期推断);若 None 则用今年。
    """
    base = os.path.basename(filename)
    stem = re.sub(r"\.(xlsx|xls)$", "", base, flags=re.IGNORECASE)

    # M_D_M_D + 产品线名
    m = re.match(r"^\s*(\d{1,2})_(\d{1,2})_(\d{1,2})_(\d{1,2})\s*(.+?)\s*$", stem)
    if not m:
        raise FilenameParseError(
            f"档名格式不符,预期 'M_D_M_D产品线.xlsx',收到:{base!r}"
        )

    m1, d1, m2, d2, product_line = m.groups()
    product_line = product_line.strip()
    if not product_line:
        raise FilenameParseError(f"档名缺产品线:{base!r}")

    if default_year is None:
        default_year = datetime.date.today().year

    try:
        start = datetime.date(default_year, int(m1), int(d1))
        # 跨年的周(例:12_28_1_3):结束月 < 起始月 → 结束年 +1
        end_year = default_year + 1 if int(m2) < int(m1) else default_year
        end = datetime.date(end_year, int(m2), int(d2))
    except ValueError as e:
        raise FilenameParseError(f"档名日期非法:{base!r} ({e})")

    return {
        "product_line": product_line,
        "week_start": start.isoformat(),
        "week_end": end.isoformat(),
    }


# ---------- 值正规化 ----------
def _to_iso_date(val):
    """Excel serial number / datetime / Timestamp / ISO 字串 -> 'YYYY-MM-DD'"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        raise MissingColumnError("日期栏出现空值")
    # pandas / datetime
    if isinstance(val, (pd.Timestamp, datetime.datetime, datetime.date)):
        return pd.Timestamp(val).date().isoformat()
    # Excel 序列号(纯数字)
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return (_EXCEL_EPOCH + datetime.timedelta(days=float(val))).date().isoformat()
    # 字串:交给 pandas 解析
    s = str(val).strip()
    try:
        return pd.to_datetime(s).date().isoformat()
    except Exception:
        raise FileCorruptError(f"无法解析日期值:{val!r}")


def _to_amount_int(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        raise MissingColumnError("金额栏出现空值")
    try:
        return int(round(float(val)))
    except (ValueError, TypeError):
        raise FileCorruptError(f"金额无法转 int:{val!r}")


def _resolve_columns(df_columns):
    """把实际表头对应到 DB 栏名;缺必要栏 -> MissingColumnError"""
    actual = {str(c).strip(): c for c in df_columns}
    mapping = {}
    for db_col, aliases in COLUMN_ALIASES.items():
        found = next((actual[a] for a in aliases if a in actual), None)
        if found is None:
            raise MissingColumnError(
                f"缺栏位 {db_col}(可接受表头之一:{aliases});"
                f"实际表头:{list(actual.keys())}"
            )
        mapping[db_col] = found
    return mapping


# ---------- 主解析 ----------
def parse_excel(filepath):
    """读 Excel -> list[dict],每 dict 含 7 个 DB 栏。"""
    try:
        df = pd.read_excel(filepath, engine="openpyxl")
    except Exception as e:
        raise FileCorruptError(f"档案无法开启 / 损坏:{e}")

    if df.empty:
        raise MissingColumnError("Excel 无资料列")

    colmap = _resolve_columns(df.columns)

    rows = []
    for idx, r in df.iterrows():
        try:
            merchant_order_no = str(r[colmap["merchant_order_no"]]).strip()
            if not merchant_order_no or merchant_order_no.lower() == "nan":
                # 空白行跳过(不算错误)
                continue
            rows.append({
                "date_iso":          _to_iso_date(r[colmap["date_iso"]]),
                "merchant_tg":       str(r[colmap["merchant_tg"]]).strip(),
                "merchant_order_no": merchant_order_no,
                "platform_order_no": str(r[colmap["platform_order_no"]]).strip(),
                "amount":            _to_amount_int(r[colmap["amount"]]),
                "payment_type":      str(r[colmap["payment_type"]]).strip(),
                "platform_id":       str(r[colmap["platform_id"]]).strip(),
            })
        except ExcelParseError as e:
            raise type(e)(f"第 {idx + 2} 列解析失败:{e}")  # +2: 表头 + 1-based

    if not rows:
        raise MissingColumnError("Excel 没有有效资料列")

    return rows


def first_date_year(rows):
    """从解析结果取第一笔日期的年份(供 parse_filename 推断年份)。"""
    if rows:
        return int(rows[0]["date_iso"][:4])
    return None
