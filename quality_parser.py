#!/usr/bin/env python3
"""
Phase 4.1 — 质检 Excel 解析模组(独立于退费 excel_parser.py)

对外:
  resolve_inspect_date(filename, file_bytes, now_iso) -> (iso_date, source)
      三层降级:档名 → docProps metadata → 上传时间
  parse_quality_excel(file_bytes) -> {sheet1_rows, sheet2_rows, sheet3_rules}

例外:
  QualityParseError — 解析失败(档名/级别/表头),上层回 400
"""
import datetime
import io
import re
import zipfile
import xml.etree.ElementTree as ET

import openpyxl


class QualityParseError(Exception):
    pass


# ════════════════════════════════════════════════════════════════
#  「审核日期」三层降级解析
# ════════════════════════════════════════════════════════════════
def _valid_date(y, m, d):
    """组日期,合法回 ISO 字串,否则 None。"""
    try:
        return datetime.date(int(y), int(m), int(d)).isoformat()
    except (ValueError, TypeError):
        return None


def parse_inspect_date_from_filename(filename):
    """
    从档名解审核日期。容忍多种格式,不硬要求「质检」前缀。
      质检_2026-06-11 / 质检 2026-06-11 / 质检_2026_06_11   → ISO 带年
      质检报告_20260611                                      → 紧凑 8 位
      质检_6月11日 / 质检 6月11日(可带 2026年)              → M月D日(年用当年)
      质检 6-11 / 质检_6_11                                   → M-D(年用当年)
    成功 → (iso_date, 'filename');失败 → None
    """
    if not filename:
        return None
    base = str(filename)
    stem = re.sub(r"\.(xlsx|xls)$", "", base, flags=re.IGNORECASE)
    this_year = datetime.date.today().year

    # 1) 带年完整日期:依序尝试结构化 pattern(由严格到宽松)
    patterns_with_year = [
        r"(20\d{2})[-_\.](\d{1,2})[-_\.](\d{1,2})",   # 2026-06-11 / 2026_06_11
        r"(20\d{2})\s+(\d{1,2})\s+(\d{1,2})",          # 2026 06 11
        r"(20\d{2})年(\d{1,2})月(\d{1,2})日",           # 2026年6月11日
        r"(20\d{2})(\d{2})(\d{2})",                     # 20260611 紧凑
    ]
    for pat in patterns_with_year:
        mm = re.search(pat, stem)
        if mm:
            iso = _valid_date(mm.group(1), mm.group(2), mm.group(3))
            if iso:
                return (iso, "filename")

    # 2) 不带年:M月D日 / M-D / M_D(年份用当年)
    patterns_no_year = [
        r"(\d{1,2})月(\d{1,2})日",     # 6月11日
        r"(?<!\d)(\d{1,2})[-_](\d{1,2})(?!\d)",   # 6-11 / 6_11(前后不接数字,避免吃到长串)
    ]
    for pat in patterns_no_year:
        mm = re.search(pat, stem)
        if mm:
            iso = _valid_date(this_year, mm.group(1), mm.group(2))
            if iso:
                return (iso, "filename")

    return None


def parse_inspect_date_from_metadata(file_bytes):
    """
    拆 .xlsx(zip)读 docProps/core.xml,取 dcterms:created / dcterms:modified
    的较早者日期作为审核日。成功 → (iso_date, 'metadata');失败 → None
    """
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
            if "docProps/core.xml" not in z.namelist():
                return None
            xml = z.read("docProps/core.xml")
    except (zipfile.BadZipFile, KeyError, OSError):
        return None

    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        return None

    ns = {"dcterms": "http://purl.org/dc/terms/"}
    dates = []
    for tag in ("created", "modified"):
        el = root.find(f"dcterms:{tag}", ns)
        if el is not None and el.text:
            # 形如 2026-06-11T08:30:00Z
            mm = re.match(r"(\d{4})-(\d{2})-(\d{2})", el.text.strip())
            if mm:
                iso = _valid_date(mm.group(1), mm.group(2), mm.group(3))
                if iso:
                    dates.append(iso)
    if not dates:
        return None
    return (min(dates), "metadata")   # 较早者


def fallback_to_upload_time(now_iso):
    """最差情境:用上传日期。source 标 'upload_time',前端可据此 warning。"""
    return (now_iso[:10], "upload_time")


def resolve_inspect_date(filename, file_bytes, now_iso):
    """三层降级主流程,保证不抛错。"""
    return (
        parse_inspect_date_from_filename(filename)
        or parse_inspect_date_from_metadata(file_bytes)
        or fallback_to_upload_time(now_iso)
    )


# ════════════════════════════════════════════════════════════════
#  Excel 三 Sheet 解析
# ════════════════════════════════════════════════════════════════
# DB 栏 -> 可接受的 Excel 表头别名(按优先序,匹配时去空白)
SHEET1_ALIASES = {
    "case_no":       ["序号", "编号", "No", "no"],
    "shift":         ["班别", "班次"],
    "agent_name":    ["客服中文名", "客服姓名", "客服名", "姓名", "客服中文"],
    "agent_account": ["客服账号", "客服帐号", "账号", "帐号", "客服ID", "客服id"],
    "case_time":     ["时间", "時間"],
    "app_name":      ["App名称", "APP名称", "app名称", "应用名称", "产品", "App", "APP"],
    "session_id":    ["会话ID", "会话id", "會話ID", "会话编号", "sessionID", "session_id"],
    "user_uid":      ["UID", "uid", "用户UID", "用户ID", "用户id"],
    "error_level":   ["错误级别", "错误等级", "级别", "等级"],
    "deduction":     ["扣分", "扣分值", "分值"],
    "error_desc":    ["错误描述", "问题描述", "描述"],
    "correct_reply": ["正确回复方式", "正确回复", "正确回覆方式", "正确话术", "建议回复"],
    "conversation":  ["完整对话上下文", "对话上下文", "完整对话", "对话内容", "上下文"],
}
SHEET1_REQUIRED = ["agent_name", "error_level", "deduction"]

SHEET2_ALIASES = {
    "shift":          ["班别", "班次"],
    "agent_name":     ["客服中文名", "客服姓名", "客服名", "姓名", "客服中文"],
    "agent_account":  ["客服账号", "客服帐号", "账号", "帐号", "客服ID", "客服id"],
    "total_messages": ["总讯息数", "总信息数", "总消息数", "讯息数", "消息数"],
    "severe_count":   ["严重(次)", "严重（次）", "严重次数", "严重"],
    "medium_count":   ["中等(次)", "中等（次）", "中等次数", "中等"],
    "minor_count":    ["轻微(次)", "轻微（次）", "轻微次数", "轻微"],
    "deduction_sum":  ["扣分总和", "扣分合计", "总扣分", "扣分"],
    "pass_rate":      ["合格率", "合格率(%)", "合格率（%）", "通过率"],
    "note":           ["备注", "備註", "说明"],
}
SHEET2_REQUIRED = ["agent_name"]

VALID_LEVELS = {"严重", "中等", "轻微"}
# 接受单一级别,也接受用 + 连接的复合级别(组长标记「一笔事件触及多条规则」)
# 例:严重 / 严重+中等 / 严重+中等+轻微。本期原样存,不拆笔、不 reconcile。
_LEVEL_RE = re.compile(r"(严重|中等|轻微)(\+(严重|中等|轻微))*")
_APP_CODE_RE = re.compile(r"[（(]\s*([A-Za-z]{2,4}-?\d+)\s*[)）]")


def _norm(s):
    return re.sub(r"\s+", "", str(s)) if s is not None else ""


def _find_header_row(rows, must_have_aliases):
    """在前 8 行里找含某必要别名的表头行,回 (header_idx, header_cells) 或 None。"""
    flat = set()
    for al in must_have_aliases:
        for a in al:
            flat.add(_norm(a))
    for i, row in enumerate(rows[:8]):
        cells = [_norm(c) for c in row]
        if any(c in flat for c in cells):
            return i, list(row)
    return None


def _map_columns(header_cells, alias_map):
    """表头 → {field: col_index}。表头去空白后,逐 field 找第一个命中的别名。"""
    norm_header = [_norm(c) for c in header_cells]
    colmap = {}
    for field, aliases in alias_map.items():
        for a in aliases:
            na = _norm(a)
            if na in norm_header:
                colmap[field] = norm_header.index(na)
                break
    return colmap


def _cell(row, colmap, field):
    idx = colmap.get(field)
    if idx is None or idx >= len(row):
        return None
    v = row[idx]
    if isinstance(v, str):
        v = v.strip()
        if v == "":
            return None
    return v


def _to_int(v, default=None):
    if v is None or v == "":
        return default
    try:
        return int(float(v))
    except (ValueError, TypeError):
        return default


def _to_float(v, default=None):
    if v is None or v == "":
        return default
    try:
        f = float(v)
        if f != f:   # NaN
            return default
        return f
    except (ValueError, TypeError):
        return default


def _parse_pass_rate(v):
    """'97.44%' / '100.00%' / 0.9744 / '0.9744' / 97.44 → 0-1 float;空/NaN → None。"""
    if v is None:
        return None
    if isinstance(v, str):
        s = v.strip()
        if s == "" or s.lower() == "nan":
            return None
        if s.endswith("%"):
            f = _to_float(s[:-1])
            return None if f is None else round(f / 100.0, 6)
        f = _to_float(s)
        if f is None:
            return None
        return round(f / 100.0, 6) if f > 1 else round(f, 6)
    f = _to_float(v)
    if f is None:
        return None
    return round(f / 100.0, 6) if f > 1 else round(f, 6)


def _extract_app_code(app_name):
    if not app_name:
        return None
    m = _APP_CODE_RE.search(str(app_name))
    return m.group(1) if m else None


def _read_sheet_rows(ws):
    """read_only 模式读出所有 row(values_only),回 list[tuple]。"""
    out = []
    for row in ws.iter_rows(values_only=True):
        out.append(row)
    return out


def parse_quality_excel(file_bytes):
    """
    用 openpyxl read_only data_only 读 3 sheet。
    回 {sheet1_rows: [...], sheet2_rows: [...], sheet3_rules: str|None}
    Sheet1/Sheet2 缺失或表头不符 → QualityParseError。Sheet3 缺失 → rules=None。
    """
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    except Exception as e:
        raise QualityParseError(f"无法开启 Excel:{e}")

    try:
        sheets = wb.worksheets
        if len(sheets) < 2:
            raise QualityParseError("Excel 至少需要 Sheet 1(错误案例总表)与 Sheet 2(客服汇总)")

        # 依名称匹配,匹配不到则退回位置(0/1/2)
        def pick(keywords, fallback_idx):
            for ws in sheets:
                t = _norm(ws.title)
                if any(k in t for k in keywords):
                    return ws
            return sheets[fallback_idx] if fallback_idx < len(sheets) else None

        ws1 = pick(["错误案例", "案例总表", "错误"], 0)
        ws2 = pick(["客服汇总", "汇总", "合格率"], 1)
        ws3 = pick(["评分标准", "评分", "标准摘要"], 2)

        rows1 = _read_sheet_rows(ws1) if ws1 is not None else []
        rows2 = _read_sheet_rows(ws2) if ws2 is not None else []

        # ---- Sheet 1 ----
        hdr1 = _find_header_row(rows1, [SHEET1_ALIASES["agent_name"], SHEET1_ALIASES["error_level"]])
        if hdr1 is None:
            raise QualityParseError("Sheet 1 找不到表头(需含「客服中文名」「错误级别」)")
        h1_idx, h1_cells = hdr1
        col1 = _map_columns(h1_cells, SHEET1_ALIASES)
        missing1 = [f for f in SHEET1_REQUIRED if f not in col1]
        if missing1:
            raise QualityParseError(f"Sheet 1 缺必要栏位:{missing1}(表头:{[_norm(c) for c in h1_cells]})")

        sheet1_rows = []
        for row in rows1[h1_idx + 1:]:
            if row is None or all(c is None or (isinstance(c, str) and c.strip() == "") for c in row):
                continue
            agent = _cell(row, col1, "agent_name")
            level = _cell(row, col1, "error_level")
            if agent is None and level is None:
                continue
            level = _norm(level)
            if not _LEVEL_RE.fullmatch(level):
                raise QualityParseError(
                    f"Sheet 1 错误级别无法识别:{level!r}(序号 {_cell(row, col1, 'case_no')},"
                    f"仅接受 严重/中等/轻微,或用 + 连接的复合级别如 严重+中等)"
                )
            ded = _to_float(_cell(row, col1, "deduction"))
            if ded is None:
                raise QualityParseError(f"Sheet 1 扣分缺失或非数字(客服 {agent},级别 {level})")
            ded = -abs(ded)   # 统一存负数
            app_name = _cell(row, col1, "app_name")
            sheet1_rows.append({
                "case_no":       _to_int(_cell(row, col1, "case_no")),
                "shift":         _cell(row, col1, "shift"),
                "agent_name":    str(agent).strip() if agent is not None else None,
                "agent_account": _cell(row, col1, "agent_account"),
                "case_time":     str(_cell(row, col1, "case_time")) if _cell(row, col1, "case_time") is not None else None,
                "app_name":      str(app_name) if app_name is not None else None,
                "app_code":      _extract_app_code(app_name),
                "session_id":    str(_cell(row, col1, "session_id")) if _cell(row, col1, "session_id") is not None else None,
                "user_uid":      str(_cell(row, col1, "user_uid")) if _cell(row, col1, "user_uid") is not None else None,
                "error_level":   level,
                "deduction":     ded,
                "error_desc":    _cell(row, col1, "error_desc"),
                "correct_reply": _cell(row, col1, "correct_reply"),
                "conversation":  _cell(row, col1, "conversation"),
            })

        if not sheet1_rows:
            raise QualityParseError("Sheet 1 没有任何有效案例资料")

        # ---- Sheet 2 ----
        hdr2 = _find_header_row(rows2, [SHEET2_ALIASES["agent_name"], SHEET2_ALIASES["pass_rate"]])
        if hdr2 is None:
            raise QualityParseError("Sheet 2 找不到表头(需含「客服中文名」「合格率」)")
        h2_idx, h2_cells = hdr2
        col2 = _map_columns(h2_cells, SHEET2_ALIASES)
        missing2 = [f for f in SHEET2_REQUIRED if f not in col2]
        if missing2:
            raise QualityParseError(f"Sheet 2 缺必要栏位:{missing2}(表头:{[_norm(c) for c in h2_cells]})")

        sheet2_rows = []
        for row in rows2[h2_idx + 1:]:
            if row is None or all(c is None or (isinstance(c, str) and c.strip() == "") for c in row):
                continue
            agent = _cell(row, col2, "agent_name")
            if agent is None:
                continue
            note = _cell(row, col2, "note")
            if isinstance(note, str) and note.strip().lower() == "nan":
                note = None
            sheet2_rows.append({
                "shift":          _cell(row, col2, "shift"),
                "agent_name":     str(agent).strip(),
                "agent_account":  _cell(row, col2, "agent_account"),
                "total_messages": _to_int(_cell(row, col2, "total_messages")),
                "severe_count":   _to_int(_cell(row, col2, "severe_count"), 0),
                "medium_count":   _to_int(_cell(row, col2, "medium_count"), 0),
                "minor_count":    _to_int(_cell(row, col2, "minor_count"), 0),
                "deduction_sum":  abs(_to_float(_cell(row, col2, "deduction_sum"), 0.0)),
                "pass_rate":      _parse_pass_rate(_cell(row, col2, "pass_rate")),
                "note":           note,
            })

        if not sheet2_rows:
            raise QualityParseError("Sheet 2 没有任何有效汇总资料")

        # ---- Sheet 3(本期只存 raw text)----
        sheet3_rules = None
        if ws3 is not None and ws3 is not ws1 and ws3 is not ws2:
            parts = []
            for row in _read_sheet_rows(ws3):
                for c in row:
                    if c is not None and str(c).strip() != "":
                        parts.append(str(c).strip())
            sheet3_rules = "\n".join(parts) if parts else None

        return {
            "sheet1_rows": sheet1_rows,
            "sheet2_rows": sheet2_rows,
            "sheet3_rules": sheet3_rules,
        }
    finally:
        wb.close()
