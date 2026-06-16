#!/usr/bin/env python3
"""
Phase 2.3 — 退费查询模组

把 /api/refunds/data 用到的查询逻辑模组化,app.py 只负责接 query string + 组 JSON。

对外:
  resolve_time_range(range_key, start, end, today=None)        -> (start_iso, end_iso, label)
  resolve_compare_range(compare_key, cur_start_iso, cur_end_iso) -> (start_iso, end_iso, label) | None
  get_kpi / get_trend / get_payment_breakdown / get_amount_tiers / get_platform_top10
  build_period(conn, start_iso, end_iso)                       -> dict(含上面 5 区块)
  compute_insights(current, compare)                          -> list[dict] (最多 3 条)
"""
import datetime

# 金额档位上界(元);> 最后一档归到 OVERFLOW
_TIER_BOUNDS = [50, 100, 150, 200, 300, 500]
_TIER_OVERFLOW = 999  # 代表 500+


# ---------- 日期工具 ----------
def _d(iso):
    return datetime.date.fromisoformat(iso)


def _md(d):
    return f"{d.month}/{d.day}"


def _quarter_start(d):
    q_month = ((d.month - 1) // 3) * 3 + 1
    return datetime.date(d.year, q_month, 1)


# ---------- 时间范围解析 ----------
def resolve_time_range(range_key, start=None, end=None, today=None):
    """range_key: week|month|quarter|year|ytd|custom -> (start_iso, end_iso, label)"""
    today = today or datetime.date.today()
    range_key = (range_key or "week").lower()

    if range_key == "week":
        s = today - datetime.timedelta(days=today.weekday())  # 本周一
        e = s + datetime.timedelta(days=6)                     # 本周日
        label = f"本周 ({_md(s)}-{_md(e)})"
    elif range_key == "month":
        s = today.replace(day=1)
        e = today
        label = f"本月 ({_md(s)}-{_md(e)})"
    elif range_key == "quarter":
        s = _quarter_start(today)
        e = today
        label = f"本季 ({_md(s)}-{_md(e)})"
    elif range_key in ("year", "ytd"):
        s = today.replace(month=1, day=1)
        e = today
        label = f"本年 ({s.year})"
    elif range_key == "custom":
        if not start or not end:
            raise ValueError("custom range 需要 start_date 与 end_date")
        s = _d(start)
        e = _d(end)
        label = f"自订 ({_md(s)}-{_md(e)})"
    else:
        raise ValueError(f"未知 range: {range_key}")

    if e < s:
        raise ValueError("end_date 早于 start_date")
    return s.isoformat(), e.isoformat(), label


def resolve_compare_range(compare_key, cur_start_iso, cur_end_iso):
    """compare_key: none|last-week|last-month|last-quarter|last-year|yoy"""
    compare_key = (compare_key or "none").lower()
    if compare_key == "none":
        return None
    cs, ce = _d(cur_start_iso), _d(cur_end_iso)

    if compare_key == "last-week":
        s = cs - datetime.timedelta(days=7)
        e = ce - datetime.timedelta(days=7)
        label = f"上周 ({_md(s)}-{_md(e)})"
    elif compare_key == "last-month":
        prev_end = cs.replace(day=1) - datetime.timedelta(days=1)  # 上月最后一天
        s = prev_end.replace(day=1)
        e = prev_end
        label = f"上月 ({_md(s)}-{_md(e)})"
    elif compare_key == "last-quarter":
        prev_q_end = _quarter_start(cs) - datetime.timedelta(days=1)
        s = _quarter_start(prev_q_end)
        e = prev_q_end
        label = f"上季 ({_md(s)}-{_md(e)})"
    elif compare_key == "last-year":
        y = cs.year - 1
        s = datetime.date(y, 1, 1)
        e = datetime.date(y, 12, 31)
        label = f"去年 ({y})"
    elif compare_key == "yoy":
        s = _shift_year(cs, -1)
        e = _shift_year(ce, -1)
        label = f"去年同期 ({_md(s)}-{_md(e)})"
    else:
        raise ValueError(f"未知 compare: {compare_key}")

    return s.isoformat(), e.isoformat(), label


def _shift_year(d, delta):
    try:
        return d.replace(year=d.year + delta)
    except ValueError:
        return d.replace(year=d.year + delta, day=28)  # 2/29 → 2/28


# ---------- 各区块查询 ----------
def get_kpi(conn, start_iso, end_iso):
    row = conn.execute(
        "SELECT COUNT(*) AS orders, COALESCE(SUM(amount),0) AS total "
        "FROM refunds WHERE date_iso BETWEEN ? AND ?",
        (start_iso, end_iso),
    ).fetchone()
    orders = row["orders"]
    total = row["total"]
    days = (_d(end_iso) - _d(start_iso)).days + 1
    return {
        "total_amount": total,
        "total_orders": orders,
        "daily_avg_amount": round(total / days) if days else 0,
        "daily_avg_orders": round(orders / days) if days else 0,
        "avg_per_order": round(total / orders) if orders else 0,
    }


def get_trend(conn, start_iso, end_iso):
    rows = conn.execute(
        "SELECT date_iso, COUNT(*) AS orders, COALESCE(SUM(amount),0) AS amount "
        "FROM refunds WHERE date_iso BETWEEN ? AND ? GROUP BY date_iso",
        (start_iso, end_iso),
    ).fetchall()
    by_date = {r["date_iso"]: r for r in rows}
    out = []
    cur, end = _d(start_iso), _d(end_iso)
    while cur <= end:
        iso = cur.isoformat()
        r = by_date.get(iso)
        out.append({
            "date": iso,
            "orders": r["orders"] if r else 0,
            "amount": r["amount"] if r else 0,
        })
        cur += datetime.timedelta(days=1)
    return out


def get_payment_breakdown(conn, start_iso, end_iso):
    rows = conn.execute(
        "SELECT payment_type AS type, COUNT(*) AS orders, COALESCE(SUM(amount),0) AS amount "
        "FROM refunds WHERE date_iso BETWEEN ? AND ? GROUP BY payment_type ORDER BY amount DESC",
        (start_iso, end_iso),
    ).fetchall()
    total_orders = sum(r["orders"] for r in rows) or 1
    return [
        {
            "type": r["type"],
            "orders": r["orders"],
            "amount": r["amount"],
            "pct_orders": round(r["orders"] / total_orders * 100, 1),
        }
        for r in rows
    ]


def get_amount_tiers(conn, start_iso, end_iso):
    rows = conn.execute(
        "SELECT amount FROM refunds WHERE date_iso BETWEEN ? AND ?",
        (start_iso, end_iso),
    ).fetchall()
    counts = {b: 0 for b in _TIER_BOUNDS}
    counts[_TIER_OVERFLOW] = 0
    for r in rows:
        amt = r["amount"]
        placed = False
        for b in _TIER_BOUNDS:
            if amt <= b:
                counts[b] += 1
                placed = True
                break
        if not placed:
            counts[_TIER_OVERFLOW] += 1
    tiers = [{"tier": b, "count": counts[b]} for b in _TIER_BOUNDS]
    tiers.append({"tier": _TIER_OVERFLOW, "count": counts[_TIER_OVERFLOW]})
    return tiers


def get_platform_top10(conn, start_iso, end_iso):
    rows = conn.execute(
        "SELECT platform_id, COUNT(*) AS orders, COALESCE(SUM(amount),0) AS amount "
        "FROM refunds WHERE date_iso BETWEEN ? AND ? "
        "GROUP BY platform_id ORDER BY amount DESC LIMIT 10",
        (start_iso, end_iso),
    ).fetchall()
    return [
        {"platform_id": r["platform_id"], "orders": r["orders"], "amount": r["amount"]}
        for r in rows
    ]


def build_period(conn, start_iso, end_iso):
    return {
        "kpi": get_kpi(conn, start_iso, end_iso),
        "trend": get_trend(conn, start_iso, end_iso),
        "payment_breakdown": get_payment_breakdown(conn, start_iso, end_iso),
        "amount_tiers": get_amount_tiers(conn, start_iso, end_iso),
        "platform_top10": get_platform_top10(conn, start_iso, end_iso),
    }


def empty_period():
    return {
        "kpi": {
            "total_amount": 0, "total_orders": 0, "daily_avg_amount": 0,
            "daily_avg_orders": 0, "avg_per_order": 0,
        },
        "trend": [],
        "payment_breakdown": [],
        "amount_tiers": [],
        "platform_top10": [],
    }


# ---------- 洞察规则 ----------
def _pct_change(cur, prev):
    if prev == 0:
        return None
    return round((cur - prev) / prev * 100, 1)


def compute_insights(current, compare):
    """最多回 3 条;有 compare 才能算变化率规则。"""
    out = []
    ck = current["kpi"]

    if compare:
        pk = compare["kpi"]
        # 规则 1:总额变化 > ±20%
        amt_chg = _pct_change(ck["total_amount"], pk["total_amount"])
        if amt_chg is not None and abs(amt_chg) > 20:
            up = amt_chg > 0
            out.append({
                "level": "bad" if up else "good",
                "icon": "▲" if up else "▼",
                "text": f"退费总额{'上升' if up else '下降'} {abs(amt_chg)}%(对比期)",
            })
        # 规则 2:金额与笔数方向背离(总额降、笔数升)
        ord_chg = _pct_change(ck["total_orders"], pk["total_orders"])
        if amt_chg is not None and ord_chg is not None and amt_chg < 0 and ord_chg > 0:
            out.append({
                "level": "neutral",
                "icon": "⚠",
                "text": "总额下降但笔数上升,小额退费比例增加",
            })
        # 规则 4:平均单笔变化 > ±15%
        avg_chg = _pct_change(ck["avg_per_order"], pk["avg_per_order"])
        if avg_chg is not None and abs(avg_chg) > 15:
            up = avg_chg > 0
            out.append({
                "level": "bad" if up else "good",
                "icon": "▲" if up else "▼",
                "text": f"单笔均额{'上升' if up else '下降'} {abs(avg_chg)}%",
            })

    # 规则 3:单一平台占金额 > 30%(只需 current)
    total_amt = ck["total_amount"]
    if total_amt > 0 and current.get("platform_top10"):
        top = current["platform_top10"][0]
        share = round(top["amount"] / total_amt * 100, 1)
        if share > 30:
            out.append({
                "level": "neutral",
                "icon": "⚠",
                "text": f"{top['platform_id']} 占金额 {share}%,集中度偏高",
            })

    # 规则 5:新平台进 Top 5(需 compare)
    if compare and current.get("platform_top10") and compare.get("platform_top10"):
        cur_top5 = [p["platform_id"] for p in current["platform_top10"][:5]]
        prev_top5 = set(p["platform_id"] for p in compare["platform_top10"][:5])
        for pid in cur_top5:
            if pid not in prev_top5:
                out.append({
                    "level": "neutral",
                    "icon": "◆",
                    "text": f"{pid} 首次进入 Top 5,可能是新增退费源",
                })
                break

    return out[:3]
