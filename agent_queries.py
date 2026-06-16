#!/usr/bin/env python3
"""
Phase 3.2 — 接线量子页查询模组(主管视角)

GET /api/agents/data 的查询逻辑。app.py 只接 query string + 组 JSON。

时间范围 / 对比解析直接复用 refund_queries(同一套 week/last-week/... 语义)。

对外:
  build_period(conn, start, end, source_filter, amap, full=True) -> dict
  empty_period(full=True)
设计:每个 period 只跑一次 SELECT 把区间内 rows 全捞进内存,
KPI / trend / 明细表 / 矩阵 / 异常 / 守门全部在 Python 端算(区间子集小,几百列)。

守门(脏数据排除,只影响 KPI 平均,不删 DB):
  - 響應時長:source='yueda' 且 response_time_sec > 3600 → 排除(疑似累计秒数误填)
  - 质量分:quality_score 为 0 / <0 / >1 → 排除(0 视为漏填,不当真实 0 分)
  - intake NULL → 不算入 active_agents
"""
import datetime
import statistics

from agent_aliases import resolve_name

# 复用退费那套时间范围解析
from refund_queries import resolve_time_range, resolve_compare_range  # noqa: F401

RESPONSE_DIRTY_THRESHOLD = 3600   # 秒;yueda > 此值视为脏数据
ESCALATION_DIRTY_THRESHOLD = 1000  # 升级工单 > 此值视为脏(果冻 4/28 = 151 亿);宽松避免误伤繁忙日
QUALITY_FLOOR = 0.6               # 异常规则 3:平均质量 < 0.6
DEFAULT_ABSENCE_DAYS = 3          # 异常规则 4:连续 N 天未提交


def _d(iso):
    return datetime.date.fromisoformat(iso)


def _md(d):
    return f"{d.month}/{d.day}"


# ---------- 守门判定 ----------
def _response_valid(source, rt):
    # Phase 3.2 微调(百川确认):yueda 響應時長字段整根不可信
    # (38% > 3600 秒,九节狼达 3.16 亿秒,疑似「累计秒数」而非「单次响应」),
    # 整源排除,平均响应只采 remote。get_kpi / get_agents_table 都走这里。
    return source == "remote" and rt is not None


def _quality_valid(q):
    return q is not None and 0 < q <= 1


def _escalation_valid(esc):
    # > 1000 视为「明显物理不可能」脏数据(果冻 4/28 = 151 亿),KPI / 异常排除
    return esc is not None and esc <= ESCALATION_DIRTY_THRESHOLD


def _avg(vals):
    return round(sum(vals) / len(vals)) if vals else None


# ---------- 取数 ----------
def _fetch_rows(conn, start_iso, end_iso, source_filter):
    sql = (
        "SELECT date_iso, agent_name, source, intake, response_time_sec, "
        "quality_score, escalation_count, shift "
        "FROM daily_reports WHERE date_iso BETWEEN ? AND ?"
    )
    params = [start_iso, end_iso]
    if source_filter in ("yueda", "remote"):
        sql += " AND source = ?"
        params.append(source_filter)
    return conn.execute(sql, params).fetchall()


# ---------- KPI ----------
def get_kpi(rows):
    total_intake = sum(r["intake"] for r in rows if r["intake"] is not None)
    total_esc = sum(r["escalation_count"] for r in rows if _escalation_valid(r["escalation_count"]))
    resp_vals = [r["response_time_sec"] for r in rows
                 if _response_valid(r["source"], r["response_time_sec"])]
    qual_vals = [r["quality_score"] for r in rows if _quality_valid(r["quality_score"])]
    return {
        "total_intake": total_intake,
        "avg_response_sec": _avg(resp_vals),
        "avg_quality_score": round(sum(qual_vals) / len(qual_vals), 4) if qual_vals else None,
        "total_escalation": total_esc,
        "active_agents": None,   # 由 build_period 用 canonical 填(别名合并后去重)
    }


# ---------- 趋势 ----------
def get_trend(rows, start_iso, end_iso, canon_of):
    by_date = {}
    for r in rows:
        d = by_date.setdefault(r["date_iso"], {"intake": 0, "agents": set()})
        if r["intake"] is not None:
            d["intake"] += r["intake"]
            d["agents"].add(canon_of(r))
    out = []
    cur, end = _d(start_iso), _d(end_iso)
    while cur <= end:
        iso = cur.isoformat()
        e = by_date.get(iso)
        out.append({
            "date": iso,
            "intake": e["intake"] if e else 0,
            "agents": len(e["agents"]) if e else 0,
        })
        cur += datetime.timedelta(days=1)
    return out


# ---------- 客服明细表(已合并别名)----------
def get_agents_table(rows, canon_of):
    groups = {}
    for r in rows:
        c = canon_of(r)
        g = groups.setdefault(c, {
            "raw_names": set(), "intake": 0, "has_intake": False, "resp": [], "qual": [],
            "esc": 0, "dates": set(), "sources": set(),
        })
        g["raw_names"].add(r["agent_name"])
        g["sources"].add(r["source"])
        g["dates"].add(r["date_iso"])
        if r["intake"] is not None:
            g["intake"] += r["intake"]
            g["has_intake"] = True       # 区分「全 NULL(没填)」与「真实 0」
        if _escalation_valid(r["escalation_count"]):
            g["esc"] += r["escalation_count"]
        if _response_valid(r["source"], r["response_time_sec"]):
            g["resp"].append(r["response_time_sec"])
        if _quality_valid(r["quality_score"]):
            g["qual"].append(r["quality_score"])

    table = []
    for c, g in groups.items():
        srcs = g["sources"]
        table.append({
            "canonical_name": c,
            "raw_names": sorted(g["raw_names"]),
            "intake": g["intake"] if g["has_intake"] else None,
            "avg_response_sec": _avg(g["resp"]),
            "avg_quality_score": round(sum(g["qual"]) / len(g["qual"]), 4) if g["qual"] else None,
            "escalation_count": g["esc"],
            "days_active": len(g["dates"]),
            "source": next(iter(srcs)) if len(srcs) == 1 else "mixed",
        })
    # intake 排序:有数据的按量降序在前,全 NULL(intake=None)排最后
    table.sort(key=lambda x: (x["intake"] is None, -(x["intake"] or 0), x["canonical_name"]))
    return table


# ---------- 热力图矩阵 ----------
def get_matrix(rows, start_iso, end_iso, canon_of):
    agents = sorted({canon_of(r) for r in rows})
    dates = []
    cur, end = _d(start_iso), _d(end_iso)
    while cur <= end:
        dates.append(cur.isoformat())
        cur += datetime.timedelta(days=1)
    cell_map = {}
    for r in rows:
        if r["intake"] is None:
            continue
        key = (canon_of(r), r["date_iso"])
        cell_map[key] = cell_map.get(key, 0) + r["intake"]
    cells = [
        {"agent": a, "date": d, "intake": v}
        for (a, d), v in sorted(cell_map.items())
    ]
    return {"agents": agents, "dates": dates, "cells": cells}


# ---------- 异常名单(4 规则,本期 vs 同期中位数)----------
# 严重度排序:质量 > 接待量 > 响应 > 不提交
_SEVERITY = {"low_quality": 0, "low_intake": 1, "slow_response": 2, "no_submit": 3}


def get_anomalies(rows, table, start_iso, end_iso, canon_of, absence_days=DEFAULT_ABSENCE_DAYS):
    if not table:
        return []

    # 中位数只算「有填 intake」的人(全 NULL 没填 ≠ 真实 0,不拉低门槛)
    intake_vals = [t["intake"] for t in table if t["intake"] is not None]
    intake_median = statistics.median(intake_vals) if intake_vals else None
    resp_vals = [t["avg_response_sec"] for t in table if t["avg_response_sec"] is not None]
    resp_median = statistics.median(resp_vals) if resp_vals else None

    # 每人在区间内有 report 的日期(判断「连续 N 天未提交」)
    dates_by_agent = {}
    for r in rows:
        dates_by_agent.setdefault(canon_of(r), set()).add(r["date_iso"])
    end = _d(end_iso)
    last_n = {(end - datetime.timedelta(days=k)).isoformat() for k in range(absence_days)}
    # 区间不足 N 天则不触发规则 4(避免短区间误判)
    span_days = (end - _d(start_iso)).days + 1

    per_agent = {}  # canonical -> 最严重一条

    def consider(agent, atype, value, threshold, note):
        cand = {"agent": agent, "type": atype, "value": value,
                "threshold": threshold, "note": note}
        cur = per_agent.get(agent)
        if cur is None or _SEVERITY[atype] < _SEVERITY[cur["type"]]:
            per_agent[agent] = cand

    for t in table:
        a = t["canonical_name"]
        # 规则 3:平均质量 < 0.6
        if t["avg_quality_score"] is not None and t["avg_quality_score"] < QUALITY_FLOOR:
            consider(a, "low_quality", t["avg_quality_score"], QUALITY_FLOOR,
                     f"平均质量 {t['avg_quality_score']} < {QUALITY_FLOOR}")
        # 规则 1:接待量 < 同期中位数 × 0.5(全 NULL 没填的人跳过,不误判)
        if intake_median is not None and t["intake"] is not None:
            thr1 = round(intake_median * 0.5)
            if t["intake"] < thr1:
                consider(a, "low_intake", t["intake"], thr1,
                         f"接待量 {t['intake']} < 同期中位数 {round(intake_median)} 的 50%({thr1})")
        # 规则 2:平均响应 > 同期中位数 × 2.0(仅 remote;yueda 响应字段整根排除)
        if resp_median is not None and t["avg_response_sec"] is not None \
                and t["source"] != "yueda":
            thr2 = round(resp_median * 2.0)
            if t["avg_response_sec"] > thr2:
                consider(a, "slow_response", t["avg_response_sec"], thr2,
                         f"平均响应 {t['avg_response_sec']}秒 > 同期中位数 {round(resp_median)}秒 的 2 倍({thr2}秒)")
        # 规则 4:最近连续 N 天未提交
        if span_days >= absence_days:
            recent = dates_by_agent.get(a, set()) & last_n
            if not recent:
                consider(a, "no_submit", 0, absence_days,
                         f"最近 {absence_days} 天({_md(end - datetime.timedelta(days=absence_days - 1))}-{_md(end)})未提交日报")

    out = list(per_agent.values())
    out.sort(key=lambda x: (_SEVERITY[x["type"]], x["agent"]))
    return out


# ---------- 守门透明化(被排除的脏数据说明)----------
def compute_data_quality_notes(rows):
    notes = []

    # 響應時長:yueda 整源排除(Phase 3.2 微调)
    yueda_resp = [r for r in rows
                  if r["source"] == "yueda" and r["response_time_sec"] is not None]
    remote_resp_n = len([r for r in rows
                         if r["source"] == "remote" and r["response_time_sec"] is not None])
    if yueda_resp:
        notes.append(
            "yueda 響應時長字段整体语义不一致(疑似「累计秒数」而非「单次响应」),"
            f"本期 KPI 平均响应仅基于 remote 来源({remote_resp_n} 笔)。yueda 明细表该字段隐藏。"
        )
        # 佐证脏的程度:列出极端值(> 3600 秒)几笔
        extreme = sorted([r for r in yueda_resp
                          if r["response_time_sec"] > RESPONSE_DIRTY_THRESHOLD],
                         key=lambda r: -r["response_time_sec"])
        if extreme:
            notes.append(
                f"  · 其中 yueda 響應時長 > {RESPONSE_DIRTY_THRESHOLD} 秒的极端值 {len(extreme)} 笔,"
                f"最严重:{extreme[0]['agent_name']} {_md(_d(extreme[0]['date_iso']))} "
                f"{extreme[0]['response_time_sec']}秒"
            )

    # 质量分被排除(0 / 越界)
    bad_q = [r for r in rows if r["quality_score"] is not None
             and not (0 < r["quality_score"] <= 1)]
    if bad_q:
        notes.append(
            f"质量分异常(0 视为漏填 / 越界 <0 或 >1)排除 {len(bad_q)} 笔,平均质量不计入"
        )

    # 升级工单脏数据排除(> 1000)
    bad_esc = sorted(
        [r for r in rows if r["escalation_count"] is not None
         and r["escalation_count"] > ESCALATION_DIRTY_THRESHOLD],
        key=lambda r: -r["escalation_count"],
    )
    if bad_esc:
        notes.append(
            f"升级工单 > {ESCALATION_DIRTY_THRESHOLD} 视为脏数据(明显物理不可能),"
            f"本期排除 {len(bad_esc)} 笔,KPI 总升级不计入"
        )
        for r in bad_esc[:5]:
            notes.append(
                f"  · {r['agent_name']} {_md(_d(r['date_iso']))} 升级工单 {r['escalation_count']} 已排除"
            )
        if len(bad_esc) > 5:
            notes.append(f"  · …另有 {len(bad_esc) - 5} 笔(详见 DB)")
    return notes


# ---------- 组装 ----------
def empty_period(full=True):
    kpi = {
        "total_intake": 0, "avg_response_sec": None, "avg_quality_score": None,
        "total_escalation": 0, "active_agents": 0,
    }
    if not full:
        return {"kpi": kpi, "trend": []}
    return {
        "kpi": kpi, "trend": [], "agents_table": [],
        "matrix": {"agents": [], "dates": [], "cells": []},
        "anomalies": [], "data_quality_notes": [],
    }


def build_period(conn, start_iso, end_iso, source_filter, amap, full=True,
                 absence_days=DEFAULT_ABSENCE_DAYS):
    rows = _fetch_rows(conn, start_iso, end_iso, source_filter)

    def canon_of(r):
        return resolve_name(r["agent_name"], r["source"], amap)

    kpi = get_kpi(rows)
    # active_agents:别名合并后、intake 非 NULL 的去重人数
    kpi["active_agents"] = len({canon_of(r) for r in rows if r["intake"] is not None})
    trend = get_trend(rows, start_iso, end_iso, canon_of)

    if not full:
        return {"kpi": kpi, "trend": trend}

    table = get_agents_table(rows, canon_of)
    return {
        "kpi": kpi,
        "trend": trend,
        "agents_table": table,
        "matrix": get_matrix(rows, start_iso, end_iso, canon_of),
        "anomalies": get_anomalies(rows, table, start_iso, end_iso, canon_of, absence_days),
        "data_quality_notes": compute_data_quality_notes(rows),
    }
