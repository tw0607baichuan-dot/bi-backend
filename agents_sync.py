#!/usr/bin/env python3
"""
Phase 3.1 — 接线量子页:坐席日报同步模组

从两个公开 Google Sheet 拉 CSV，清洗，自动补日期(解法 A)，INSERT OR REPLACE 进
daily_reports 表。被 app.py 的 POST /api/agents/sync 与 cron 调用。

两个数据源:
  悦达日报  1doavB8kz97Y9T2sbUTYd5k9sFf9q7bJhf_TfGMK_B9w  3 个 tab(早/中/晚)
  远程日报  1B2bClcMntAWu89EAscHXx7SUt3RNQyaxXl0ge_piItw  单 tab(gid 0)

对外:
  sync_all(conn) -> dict   主流程,被 endpoint 调用
"""
import csv
import datetime
import io
import json
import logging
import os
import re

import requests

# ---------- 数据源设定 ----------
YUEDA_SHEET_ID = "1doavB8kz97Y9T2sbUTYd5k9sFf9q7bJhf_TfGMK_B9w"
YUEDA_GIDS = {            # shift 标签 -> gid
    "早": "131357270",
    "中": "165356656",
    "晚": "1809068869",
}
REMOTE_SHEET_ID = "1B2bClcMntAWu89EAscHXx7SUt3RNQyaxXl0ge_piItw"
REMOTE_GID = "0"

CSV_URL = "https://docs.google.com/spreadsheets/d/{sid}/export?format=csv&gid={gid}"
FETCH_TIMEOUT = 30

MIN_DATE = "2025-01-01"
MAX_DATE = f"{datetime.date.today().year}-12-31"

_EXCEL_EPOCH = datetime.datetime(1899, 12, 30)

# ---------- log ----------
log = logging.getLogger("agents_sync")
if not log.handlers:
    log.setLevel(logging.INFO)
    _fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    try:
        _fh = logging.FileHandler("/var/log/bi-backend/agents_sync.log")
        _fh.setFormatter(_fmt)
        log.addHandler(_fh)
    except OSError:
        _sh = logging.StreamHandler()
        _sh.setFormatter(_fmt)
        log.addHandler(_sh)


class SyncError(Exception):
    """拉取 / 同步失败,上层 raise → Flask 回 5xx。"""


# ---------- 拉取 ----------
def fetch_csv(sheet_id, gid):
    """拉公开 CSV → list[dict]。防呆:拉到 HTML(未公开)就 raise。"""
    url = CSV_URL.format(sid=sheet_id, gid=gid)
    try:
        resp = requests.get(url, timeout=FETCH_TIMEOUT)
        resp.raise_for_status()
    except requests.RequestException as e:
        raise SyncError(f"拉取失败 sheet={sheet_id} gid={gid}: {e}")
    resp.encoding = "utf-8"
    text = resp.text
    head = text.lstrip()[:64].lower()
    if head.startswith("<!doctype") or head.startswith("<html"):
        raise SyncError(
            f"sheet={sheet_id} gid={gid} 拉到 HTML 而非 CSV"
            "(Sheet 可能未设为「知道连结的任何人 — 检视者」)"
        )
    # 原始訊息/原文 字段有内嵌换行 → 必须用 csv.DictReader,不能 split
    return list(csv.DictReader(io.StringIO(text)))


# ---------- 值清洗 ----------
def _first_number(value, signed=False):
    """取字串里第一个数字(escalation / intake 用,容忍 '2/2' '1/5' '163/163')。"""
    s = str(value if value is not None else "").strip()
    if not s or s.lower() == "nan":
        return None
    # 只看 '/' 前那段(接待量 163/163、工单 1/5 都取前数)
    head = s.split("/")[0].strip()
    m = re.search(r"-?\d+" if signed else r"\d+", head)
    return int(m.group()) if m else None


def parse_intake(value):
    """'163/163'→163, '-/180'→None, ''→None, '133'→133"""
    return _first_number(value)


def parse_escalation(value):
    """'2/2'→2, '1/5'→1, '30'→30, ''→None"""
    return _first_number(value)


def parse_response_time(value):
    """'70s'→70, '60秒'→60, '80'→80, ''→None"""
    s = str(value if value is not None else "").strip()
    if not s or s.lower() == "nan":
        return None
    m = re.search(r"\d+", s)
    return int(m.group()) if m else None


def parse_quality(value):
    """归一到 0-1 float。'90%'→0.9, '90'→0.9, '0'→0.0, '30/30'→1.0, ''→None"""
    s = str(value if value is not None else "").strip().replace("%", "")
    if not s or s.lower() == "nan":
        return None
    try:
        if "/" in s:
            num, den = s.split("/", 1)
            num, den = float(num), float(den)
            if den == 0:
                return None
            v = num / den
        else:
            v = float(s)
    except ValueError:
        return None
    if v > 1:            # 视为百分制 0-100
        v = v / 100.0
    if v < 0:
        return None
    return round(v, 4)


def parse_date(value):
    """ISO / 美式 m/d/Y / Excel serial → 'YYYY-MM-DD';失败回 None。"""
    if value is None:
        return None
    s = str(value).strip()
    if not s or s.lower() == "nan":
        return None
    m = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})", s)          # ISO
    if m:
        y, mo, d = map(int, m.groups())
    else:
        m = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{4})", s)      # 美式
        if m:
            mo, d, y = map(int, m.groups())
        elif re.match(r"^\d+(\.\d+)?$", s):                         # Excel serial
            try:
                return (_EXCEL_EPOCH + datetime.timedelta(days=float(s))).date().isoformat()
            except (ValueError, OverflowError):
                return None
        else:
            return None
    try:
        return datetime.date(y, mo, d).isoformat()
    except ValueError:
        return None


def parse_record_time(value):
    """收錄時間原样存(仅 strip)。"""
    s = str(value if value is not None else "").strip()
    return s or None


def record_date(value):
    """从收錄時間抽 ISO 日期(用来补日期);非可解析格式回 None。"""
    s = str(value if value is not None else "").strip()
    m = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})", s)
    if not m:
        return None
    y, mo, d = map(int, m.groups())
    try:
        return datetime.date(y, mo, d).isoformat()
    except ValueError:
        return None


# ---------- 载入两个数据源 ----------
def _row(date_iso, agent, source, **kw):
    base = {
        "date_iso": date_iso,
        "original_date_iso": None,
        "correction_note": None,
        "agent_name": agent,
        "intake": None,
        "response_time_sec": None,
        "quality_score": None,
        "escalation_count": None,
        "shift": None,
        "source": source,
        "submit_time": None,
        "record_time": None,
        "raw_data": None,
    }
    base.update(kw)
    return base


def load_yueda():
    """悦达 3 班合并。date_iso=日期, agent_name=提交人, shift=sheet 名(早/中/晚)。"""
    out = []
    for shift, gid in YUEDA_GIDS.items():
        for raw in fetch_csv(YUEDA_SHEET_ID, gid):
            date_iso = parse_date(raw.get("日期"))
            agent = (raw.get("提交人") or "").strip()
            if not date_iso or not agent:
                continue
            out.append(_row(
                date_iso, agent, "yueda",
                intake=parse_intake(raw.get("接待/處理")),
                response_time_sec=parse_response_time(raw.get("響應時長(秒)")),
                quality_score=parse_quality(raw.get("質量指標")),
                escalation_count=parse_escalation(raw.get("升級工單")),
                shift=shift,
                submit_time=(raw.get("提交時間") or "").strip() or None,
                raw_data=json.dumps(raw, ensure_ascii=False),
            ))
    return out


def load_remote():
    """远程日报单 tab。date_iso=日期, agent_name=提交人, record_time=收錄時間。"""
    out = []
    for raw in fetch_csv(REMOTE_SHEET_ID, REMOTE_GID):
        date_iso = parse_date(raw.get("日期"))
        agent = (raw.get("提交人") or "").strip()
        if not date_iso or not agent:
            continue
        shift_raw = (raw.get("班次") or "").strip()
        shift = "白" if "白" in shift_raw else ("夜" if "夜" in shift_raw else (shift_raw or None))
        out.append(_row(
            date_iso, agent, "remote",
            intake=parse_intake(raw.get("接待量")),
            response_time_sec=parse_response_time(raw.get("響應時長")),
            quality_score=parse_quality(raw.get("質量指標")),
            escalation_count=parse_escalation(raw.get("升級工單數")),
            shift=shift,
            record_time=parse_record_time(raw.get("收錄時間")),
            raw_data=json.dumps(raw, ensure_ascii=False),
        ))
    return out


# ---------- 自动补日期(解法 A,仅 remote)----------
def _rt_key(r):
    """排序键:record_time 升序,None 排最后。"""
    return (r["record_time"] is None, r["record_time"] or "")


def fix_overpaste_dates(rows):
    """
    解法 A:同一人同一天多笔(抢贴贴错日期),用收錄時間推真实日期。
      - 每 (agent, original_date) group 按 record_time 升序
      - 第 1 笔保留原日期;第 N(>=2)笔取收錄時間的日期当 date_iso
      - record_time 缺失/无法解析 / 推算日期 == 原始 / 超出合理范围 → 不修(log)
      - 修正后仍撞 (date_iso, agent) → 保留 record_time 较早那笔,其余 skip(log)
    回 (kept_rows, corrections, skipped, warnings)。
    """
    corrections, warnings = [], []

    groups = {}
    for r in rows:
        groups.setdefault((r["agent_name"], r["date_iso"]), []).append(r)

    for (agent, orig), grp in groups.items():
        if len(grp) < 2:
            continue
        grp.sort(key=_rt_key)
        for r in grp[1:]:
            rt = r["record_time"]
            new = record_date(rt)
            if new is None:
                warnings.append(f"{agent} {orig}: 收錄時間缺失/无法解析({rt!r}),不修,保留原日期")
                continue
            if new < MIN_DATE or new > MAX_DATE:
                warnings.append(f"{agent} {orig}: 推算日期 {new} 超出合理范围 [{MIN_DATE}~{MAX_DATE}],不修")
                continue
            if new == orig:
                warnings.append(f"{agent} {orig}: 收錄時间日期与原始相同,解法 A 无法分离,保留原日期(交去重处理)")
                continue
            r["original_date_iso"] = orig
            r["date_iso"] = new
            r["correction_note"] = f"原始抢贴 {orig} → 修正为 {new}"
            corrections.append({
                "agent": agent, "original": orig, "corrected": new, "record_time": rt,
            })

    # 边界:第 1 笔收錄日与填写日相差 >3 天 → 仅记录不修
    for (agent, orig), grp in groups.items():
        first = min(grp, key=_rt_key)
        rd = record_date(first["record_time"])
        if rd:
            try:
                delta = abs((datetime.date.fromisoformat(rd) - datetime.date.fromisoformat(orig)).days)
                if delta > 3:
                    warnings.append(f"{agent} {orig}: 第一笔收錄日 {rd} 与填写日相差 {delta} 天(>3),仅记录不修")
            except ValueError:
                pass

    # 守门:同日多笔被去重前,比较关键字段;不一致 → WARN(只记录,不改去重结果)
    _KEY_FIELDS = ("intake", "response_time_sec", "quality_score", "escalation_count")
    dedupe_groups = {}
    for r in rows:
        dedupe_groups.setdefault((r["date_iso"], r["agent_name"]), []).append(r)
    for (d, agent), grp in dedupe_groups.items():
        if len(grp) < 2:
            continue
        mismatched = [f for f in _KEY_FIELDS if len({r[f] for r in grp}) > 1]
        if mismatched:
            detail = "; ".join(
                f"{f}=[" + ", ".join(repr(r[f]) for r in sorted(grp, key=_rt_key)) + "]"
                for f in mismatched
            )
            warnings.append(
                f"{agent} {d}: 同日 {len(grp)} 笔去重,关键字段不一致(将丢弃较晚笔)→ {detail}"
            )
            log.warning("DEDUPE-MISMATCH %s %s (%d 笔) %s", agent, d, len(grp), detail)

    # 去重:同 source 内 (date_iso, agent) 撞 → 留 record_time 最早那笔
    seen, kept, skipped = set(), [], []
    for r in sorted(rows, key=_rt_key):
        key = (r["date_iso"], r["agent_name"])
        if key in seen:
            skipped.append({
                "agent": r["agent_name"], "date": r["date_iso"], "record_time": r["record_time"],
            })
            warnings.append(f"{r['agent_name']} {r['date_iso']}: 同键重复(收錄 {r['record_time']}),skip 保留较早那笔")
        else:
            seen.add(key)
            kept.append(r)

    return kept, corrections, skipped, warnings


# ---------- 写入 ----------
_INSERT = """
INSERT OR REPLACE INTO daily_reports
    (date_iso, original_date_iso, correction_note, agent_name, intake,
     response_time_sec, quality_score, escalation_count, shift, source,
     submit_time, record_time, raw_data, synced_at)
VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
"""


def _upsert(conn, rows, synced_at):
    conn.executemany(_INSERT, [
        (
            r["date_iso"], r["original_date_iso"], r["correction_note"], r["agent_name"],
            r["intake"], r["response_time_sec"], r["quality_score"], r["escalation_count"],
            r["shift"], r["source"], r["submit_time"], r["record_time"], r["raw_data"],
            synced_at,
        )
        for r in rows
    ])
    return len(rows)


# ---------- 主流程 ----------
def sync_all(conn):
    """拉两个 Sheet → 清洗 → 补日期(remote)→ INSERT OR REPLACE。回统计 dict。"""
    synced_at = datetime.datetime.now().isoformat()
    log.info("=== sync_all 开始 %s ===", synced_at)

    yueda = load_yueda()
    remote_raw = load_remote()
    remote, corrections, skipped, warnings = fix_overpaste_dates(remote_raw)

    for w in warnings:
        log.info("WARN %s", w)
    for c in corrections:
        log.info("CORRECT %s %s → %s (收錄 %s)", c["agent"], c["original"], c["corrected"], c["record_time"])

    yueda_imported = _upsert(conn, yueda, synced_at)
    remote_imported = _upsert(conn, remote, synced_at)
    conn.commit()

    result = {
        "yueda_imported": yueda_imported,
        "yueda_skipped": 0,
        "remote_imported": remote_imported,
        "remote_skipped": len(skipped),
        "remote_corrected": len(corrections),
        "corrections": corrections,
        "warnings": warnings,
        "total": yueda_imported + remote_imported,
        "synced_at": synced_at,
    }
    log.info("=== sync_all 完成 yueda=%d remote=%d corrected=%d skipped=%d total=%d ===",
             yueda_imported, remote_imported, len(corrections), len(skipped), result["total"])
    return result


if __name__ == "__main__":
    import sqlite3
    _conn = sqlite3.connect("/var/data/refunds/db.sqlite")
    try:
        out = sync_all(_conn)
    finally:
        _conn.close()
    out_print = dict(out)
    print(json.dumps(out_print, ensure_ascii=False, indent=2))
