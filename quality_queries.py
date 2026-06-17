#!/usr/bin/env python3
"""
Phase 4.2 — 质检查询模组(主管视角)

GET /api/quality/data 的查询逻辑。app.py 只接 query string + 组 JSON。
时间范围 / 对比解析复用 refund_queries(同一套 week/last-week/... 语义,额外补 today)。

设计原则(百川指示):parser 老实存 Sheet 1 明细 + Sheet 2 汇总两套数字,不 reconcile。
  - KPI 严重/中等/轻微次数、合格率、总讯息 → 取 Sheet 2(quality_summary,组长定夺过的真值)
  - 错误案例明细、被审次数、错误类型 → 取 Sheet 1(quality_inspections,逐笔)
  「严重+中等」这种复合级别在 Sheet 1 原样存一笔,Sheet 2 拆成 severe+1/medium+1,两边各自呈现。

对外:
  resolve_quality_range(range_key, start, end)   -> (start_iso, end_iso, label)  (含 today)
  resolve_compare_range(...)                     (直接复用 refund_queries)
  build_period(conn, start, end, dept_filter, amap, full=True, agent_filter=None) -> dict
  get_improvements(conn, cur_start, cur_end, cmp_start, cmp_end, dept_filter, amap, agent_filter=None)
  empty_period(full=True)
  get_latest_rules(conn, raw_dir)                -> dict(Sheet 3 原文摘要)

supersede 过滤:所有查询一律只看 status='active' 的 upload(沿用退费 _ACTIVE_FILTER 模式)。
"""
import datetime
import os
import re

from agent_aliases import resolve_name
from refund_queries import resolve_time_range, resolve_compare_range  # noqa: F401

IMPROVE_THRESHOLD = 0.05   # 合格率变化 > 此值才进进步/退步榜
_TAG_RE = re.compile(r"^【([^】]+)】")

# 只统计 active upload;supersede 只改 status、不删明细,故查询层须自行过滤
_ACTIVE_INSPECT = (
    "AND upload_id IN (SELECT id FROM quality_uploads WHERE status = 'active')"
)
_ACTIVE_SUMMARY = _ACTIVE_INSPECT


# ---------- 日期工具 ----------
def _d(iso):
    return datetime.date.fromisoformat(iso)


def _md(d):
    return f"{d.month}/{d.day}"


def resolve_quality_range(range_key, start=None, end=None, today=None):
    """在退费那套之上补 'today'(质检按日审,常看单日)。"""
    if (range_key or "").lower() == "today":
        t = today or datetime.date.today()
        return t.isoformat(), t.isoformat(), f"今日 ({_md(t)})"
    return resolve_time_range(range_key, start, end, today=today)


# ---------- 级别工具 ----------
def _level_rank(level):
    """复合级别取最高严重度排序:严重 < 中等 < 轻微。"""
    s = level or ""
    if "严重" in s:
        return 0
    if "中等" in s:
        return 1
    if "轻微" in s:
        return 2
    return 8


def _has_severe(level):
    return "严重" in (level or "")


# ---------- 取数 ----------
def _dept_clause(dept_filter):
    if dept_filter in ("dx", "df"):
        return " AND dept = ?", [dept_filter]
    return "", []


def _fetch_summary(conn, start_iso, end_iso, dept_filter):
    clause, extra = _dept_clause(dept_filter)
    sql = (
        "SELECT inspect_date, dept, shift, agent_name, agent_account, "
        "total_messages, severe_count, medium_count, minor_count, "
        "deduction_sum, pass_rate, note "
        f"FROM quality_summary WHERE inspect_date BETWEEN ? AND ? {_ACTIVE_SUMMARY}{clause}"
    )
    return conn.execute(sql, [start_iso, end_iso] + extra).fetchall()


def _fetch_inspections(conn, start_iso, end_iso, dept_filter):
    clause, extra = _dept_clause(dept_filter)
    sql = (
        "SELECT inspect_date, dept, case_no, shift, agent_name, agent_account, "
        "case_time, app_name, app_code, session_id, user_uid, "
        "error_level, deduction, error_desc, correct_reply, conversation "
        f"FROM quality_inspections WHERE inspect_date BETWEEN ? AND ? {_ACTIVE_INSPECT}{clause}"
    )
    return conn.execute(sql, [start_iso, end_iso] + extra).fetchall()


def _match_agent(r, agent_filter):
    """agent_filter(英文 username)比对 account / 中文名(给组员自查用)。"""
    if not agent_filter:
        return True
    f = agent_filter.strip().lower()
    acc = (r["agent_account"] or "").strip().lower()
    name = (r["agent_name"] or "").strip().lower()
    return f == acc or f == name


# ---------- 汇总聚合(别名合并 + 多日聚合)----------
def _agg_summary(summary_rows, canon_of, inspect_canon=None):
    """quality_summary 按 canonical 聚合(多日:次数 sum,合格率讯息加权平均)。"""
    groups = {}
    for r in summary_rows:
        c = canon_of(r["agent_name"])
        g = groups.setdefault(c, {
            "raw_names": set(), "shifts": set(),
            "total_messages": 0, "severe": 0, "medium": 0, "minor": 0,
            "deduction_sum": 0.0, "pr_weighted": 0.0, "pr_weight": 0,
            "pr_simple": [], "notes": set(),
        })
        g["raw_names"].add(r["agent_name"])
        if r["shift"]:
            g["shifts"].add(r["shift"])
        msg = r["total_messages"] or 0
        g["total_messages"] += msg
        g["severe"] += r["severe_count"] or 0
        g["medium"] += r["medium_count"] or 0
        g["minor"] += r["minor_count"] or 0
        g["deduction_sum"] += r["deduction_sum"] or 0.0
        if r["pass_rate"] is not None:
            g["pr_simple"].append(r["pass_rate"])
            if msg > 0:
                g["pr_weighted"] += r["pass_rate"] * msg
                g["pr_weight"] += msg
        if r["note"]:
            g["notes"].add(r["note"])
    out = {}
    for c, g in groups.items():
        if g["pr_weight"] > 0:
            pr = round(g["pr_weighted"] / g["pr_weight"], 4)
        elif g["pr_simple"]:
            pr = round(sum(g["pr_simple"]) / len(g["pr_simple"]), 4)
        else:
            pr = None
        shifts = sorted(g["shifts"])
        out[c] = {
            "canonical_name": c,
            "raw_names": sorted(g["raw_names"]),
            "shift": shifts[0] if len(shifts) == 1 else ("mixed" if shifts else None),
            "total_messages": g["total_messages"],
            "severe_count": g["severe"],
            "medium_count": g["medium"],
            "minor_count": g["minor"],
            "deduction_sum": round(g["deduction_sum"], 2),
            "pass_rate": pr,
            "note": "；".join(sorted(g["notes"])) if g["notes"] else None,
            "has_cases": bool(inspect_canon and c in inspect_canon),
        }
    return out


# ---------- KPI(5 张,主管视角)----------
def get_kpi(summary_agg, inspection_rows, canon_of):
    prs = [v["pass_rate"] for v in summary_agg.values() if v["pass_rate"] is not None]
    avg_pass_rate = round(sum(prs) / len(prs), 4) if prs else None

    severe_total = sum(v["severe_count"] for v in summary_agg.values())
    total_agents = len(summary_agg)
    full_pass = sum(1 for v in summary_agg.values()
                    if v["pass_rate"] is not None and v["pass_rate"] >= 0.9999)
    total_messages = sum(v["total_messages"] for v in summary_agg.values())

    # 被审次数最高的组员(Sheet 1 逐笔)
    by_agent = {}
    for r in inspection_rows:
        c = canon_of(r["agent_name"])
        a = by_agent.setdefault(c, {"agent_name": c, "total_cases": 0, "severe_cases": 0})
        a["total_cases"] += 1
        if _has_severe(r["error_level"]):
            a["severe_cases"] += 1
    top = None
    if by_agent:
        top = sorted(by_agent.values(),
                     key=lambda x: (-x["total_cases"], -x["severe_cases"], x["agent_name"]))[0]

    return {
        "avg_pass_rate": avg_pass_rate,
        "severe_count": severe_total,
        "full_pass_agents": full_pass,
        "total_agents": total_agents,
        "total_messages_inspected": total_messages,
        "top_attention_agent": top,
    }


# ---------- 明细表(合格率升序,差的在最上)----------
def get_summary_table(summary_agg):
    table = list(summary_agg.values())
    # pass_rate 升序;None(没汇总到)排最后
    table.sort(key=lambda x: (x["pass_rate"] is None, x["pass_rate"] if x["pass_rate"] is not None else 1.0,
                              x["canonical_name"]))
    return table


# ---------- 错误案例(按组员分组)----------
def get_cases(inspection_rows, canon_of):
    groups = {}
    for r in inspection_rows:
        c = canon_of(r["agent_name"])
        g = groups.setdefault(c, {"agent_name": c, "total_cases": 0, "severe_cases": 0, "cases": []})
        g["total_cases"] += 1
        if _has_severe(r["error_level"]):
            g["severe_cases"] += 1
        g["cases"].append({
            "case_no": r["case_no"],
            "case_time": r["case_time"],
            "app_name": r["app_name"],
            "app_code": r["app_code"],
            "error_level": r["error_level"],
            "deduction": r["deduction"],
            "error_desc": r["error_desc"],
            "correct_reply": r["correct_reply"],
            "conversation": r["conversation"],
        })
    out = list(groups.values())
    # 组内:级别(严重→中等→轻微)+ case_no
    for g in out:
        g["cases"].sort(key=lambda c: (_level_rank(c["error_level"]),
                                       c["case_no"] if c["case_no"] is not None else 1e9))
    # 组间:严重数降序(优先看错最多的人)
    out.sort(key=lambda g: (-g["severe_cases"], -g["total_cases"], g["agent_name"]))
    return out


# ---------- 错误类型分布 ----------
def get_error_dist(inspection_rows, top_n=None):
    """优先抓 error_desc 开头【xx】标签;覆盖率不足一半则 fallback 用 error_level 分布。"""
    n = len(inspection_rows)
    tagged = sum(1 for r in inspection_rows if _TAG_RE.match((r["error_desc"] or "").strip()))
    use_tags = n > 0 and tagged / n >= 0.5

    dist = {}
    for r in inspection_rows:
        if use_tags:
            m = _TAG_RE.match((r["error_desc"] or "").strip())
            key = m.group(1) if m else (r["error_level"] or "未分类")
        else:
            key = r["error_level"] or "未分类"
        e = dist.setdefault(key, {"error_type": key, "count": 0, "level": r["error_level"]})
        e["count"] += 1
    items = sorted(dist.values(), key=lambda x: (-x["count"], x["error_type"]))
    if top_n:
        items = items[:top_n]
    return items


# ---------- 进步榜 / 退步榜(本期 vs 上期)----------
def get_improvements(conn, cur_start, cur_end, cmp_start, cmp_end, dept_filter, amap,
                     agent_filter=None):
    cur_rows = [r for r in _fetch_summary(conn, cur_start, cur_end, dept_filter)
                if _match_agent(r, agent_filter)]
    cmp_rows = [r for r in _fetch_summary(conn, cmp_start, cmp_end, dept_filter)
                if _match_agent(r, agent_filter)]
    if not cmp_rows:
        return None   # 对比期没资料 → 无榜

    def canon_of(name):
        return resolve_name(name, None, amap)

    cur_agg = _agg_summary(cur_rows, canon_of)
    cmp_agg = _agg_summary(cmp_rows, canon_of)

    improved, declined = [], []
    for c, cur in cur_agg.items():
        prev = cmp_agg.get(c)
        if prev is None or cur["pass_rate"] is None or prev["pass_rate"] is None:
            continue
        delta = round(cur["pass_rate"] - prev["pass_rate"], 4)
        row = {"agent_name": c, "prev_pass_rate": prev["pass_rate"],
               "cur_pass_rate": cur["pass_rate"], "delta": delta}
        if delta > IMPROVE_THRESHOLD:
            improved.append(row)
        elif delta < -IMPROVE_THRESHOLD:
            declined.append(row)
    improved.sort(key=lambda x: -x["delta"])
    declined.sort(key=lambda x: x["delta"])
    return {"improved": improved[:5], "declined": declined[:5]}


# ---------- 守门说明 ----------
def compute_data_quality_notes(summary_rows, inspection_rows):
    notes = []
    dates = sorted({r["inspect_date"] for r in summary_rows} |
                   {r["inspect_date"] for r in inspection_rows})
    if dates:
        notes.append(f"本期覆盖 {len(dates)} 个质检日 (Excel):{', '.join(dates)}")

    # 复合错误级别(组长「一笔事件触及多条规则」)— Sheet 1 与 Sheet 2 扣分各自存,未 reconcile
    compound = [r for r in inspection_rows if "+" in (r["error_level"] or "")]
    if compound:
        sample = "、".join(
            f"{r['agent_name']}(序号{r['case_no']} {r['error_level']})" for r in compound[:5]
        )
        notes.append(
            f"错误级别含复合标记(组长标记「一笔事件触及多条规则」){len(compound)} 笔:{sample}。"
            "Sheet 1 明细扣分与 Sheet 2 汇总次数/扣分各自老实存,未做 reconcile"
            "(Sheet 2 为组长最终汇总真值)。"
        )

    # 错误类型分布的呈现口径(标签覆盖率不足则 fallback)
    n = len(inspection_rows)
    tagged = sum(1 for r in inspection_rows if _TAG_RE.match((r["error_desc"] or "").strip()))
    if n > 0 and tagged / n < 0.5:
        notes.append(
            f"错误类型分布:本期 error_desc 多为自由文本(仅 {tagged}/{n} 笔有【级别】标签),"
            "暂以错误级别(严重/中等/轻微)分布呈现。"
        )
    return notes


# ---------- 组装 ----------
def empty_period(full=True):
    kpi = {
        "avg_pass_rate": None, "severe_count": 0,
        "full_pass_agents": 0, "total_agents": 0,
        "total_messages_inspected": 0, "top_attention_agent": None,
    }
    if not full:
        return {"kpi": kpi}
    return {
        "kpi": kpi, "summary_table": [], "cases": [],
        "error_type_dist": [], "improvements": None, "data_quality_notes": [],
    }


def build_period(conn, start_iso, end_iso, dept_filter, amap, full=True, agent_filter=None):
    summary_rows = [r for r in _fetch_summary(conn, start_iso, end_iso, dept_filter)
                    if _match_agent(r, agent_filter)]
    inspection_rows = [r for r in _fetch_inspections(conn, start_iso, end_iso, dept_filter)
                       if _match_agent(r, agent_filter)]

    def canon_of(name):
        return resolve_name(name, None, amap)

    inspect_canon = {canon_of(r["agent_name"]) for r in inspection_rows}
    summary_agg = _agg_summary(summary_rows, canon_of, inspect_canon=inspect_canon)
    kpi = get_kpi(summary_agg, inspection_rows, canon_of)

    if not full:
        return {"kpi": kpi}

    return {
        "kpi": kpi,
        "summary_table": get_summary_table(summary_agg),
        "cases": get_cases(inspection_rows, canon_of),
        "error_type_dist": get_error_dist(inspection_rows),
        "improvements": None,   # 由 endpoint 在有 compare 时填
        "data_quality_notes": compute_data_quality_notes(summary_rows, inspection_rows),
    }


# ---------- Sheet 3 评分标准原文(从最新 active upload 的原档现读现解)----------
def get_latest_rules(conn, raw_dir):
    """quality_uploads 没存 Sheet 3 原文,改从最新 active upload 的原档重新解析抽出。"""
    import quality_parser
    row = conn.execute(
        "SELECT id, original_filename, stored_filename, inspect_date, uploaded_at "
        "FROM quality_uploads WHERE status = 'active' "
        "ORDER BY inspect_date DESC, id DESC LIMIT 1"
    ).fetchone()
    if row is None:
        return {"has_rules": False, "message": "目前没有 active 质检上传"}
    path = os.path.join(raw_dir, row["stored_filename"] or "")
    if not row["stored_filename"] or not os.path.exists(path):
        return {
            "has_rules": False,
            "upload_id": row["id"],
            "message": f"原档不在({row['stored_filename']}),无法抽取 Sheet 3",
        }
    try:
        with open(path, "rb") as f:
            parsed = quality_parser.parse_quality_excel(f.read())
    except Exception as e:  # noqa: BLE001
        return {"has_rules": False, "upload_id": row["id"],
                "message": f"重新解析失败:{e}"}
    rules = parsed.get("sheet3_rules")
    return {
        "has_rules": rules is not None,
        "upload_id": row["id"],
        "source_file": row["original_filename"],
        "inspect_date": row["inspect_date"],
        "rules_text": rules,
    }
