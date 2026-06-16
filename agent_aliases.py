#!/usr/bin/env python3
"""
Phase 3.2 — 人名合并(别名归一)模组

daily_reports 的原始 agent_name 不动;agent_aliases 记录「别名 → 主名」。
查询层用 build_alias_map() 一次建 dict,resolve_name() 归一(不每行查 DB)。

对外:
  levenshtein_distance(a, b)
  is_prefix_match(a, b)
  sound_similarity(a, b)            -> (bool, pinyin_str|None)
  build_alias_map(conn)             -> dict 供 resolve_name 用
  resolve_name(name, source, amap)  -> canonical 显示名
  find_alias_suggestions(conn)      -> list[dict] (confidence 降序)
  list_aliases(conn, page, size)    -> dict(items,total,page,page_size)
  add_alias(conn, ...)              -> dict(新列)
  remove_alias(conn, alias_id)      -> bool
"""
import datetime

# pypinyin 可选:装得上就做同音字判断,装不上 fallback 只看 Levenshtein/前缀
try:
    from pypinyin import lazy_pinyin
    _HAS_PINYIN = True
except ImportError:
    _HAS_PINYIN = False


# ---------- fuzzy 基础 ----------
def levenshtein_distance(a, b):
    """标准 DP 编辑距离(纯 Python,不依赖外部套件)。"""
    a = a or ""
    b = b or ""
    if a == b:
        return 0
    if not a:
        return len(b)
    if not b:
        return len(a)
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(
                prev[j] + 1,        # 删除
                cur[j - 1] + 1,     # 插入
                prev[j - 1] + (ca != cb),  # 替换
            ))
        prev = cur
    return prev[-1]


def is_prefix_match(a, b):
    """a 是 b 的前缀,或 b 是 a 的前缀(且不相等)。"""
    if a == b:
        return False
    return a.startswith(b) or b.startswith(a)


def sound_similarity(a, b):
    """同音字判断:无声调拼音序列相等 → (True, 'ai-wa');不可用/不同 → (False, None)。"""
    if not _HAS_PINYIN:
        return False, None
    pa = lazy_pinyin(a)
    pb = lazy_pinyin(b)
    if pa == pb:
        return True, "-".join(pa)
    return False, None


# ---------- alias 归一(查询层用)----------
def build_alias_map(conn):
    """
    建两层 dict:
      ('别名', 'yueda'|'remote') -> 主名   (来源限定)
      ('别名', None)             -> 主名   (跨源)
    resolve_name 先查来源限定,再查跨源,都没有就用原名。
    """
    amap = {}
    for r in conn.execute(
        "SELECT canonical_name, alias_name, source FROM agent_aliases"
    ).fetchall():
        amap[(r["alias_name"], r["source"])] = r["canonical_name"]
    return amap


def resolve_name(name, source, amap):
    if (name, source) in amap:
        return amap[(name, source)]
    if (name, None) in amap:
        return amap[(name, None)]
    return name


# ---------- 配对建议 ----------
def _confidence_and_reason(a, b):
    """回 (confidence, reason) 或 None(无关系)。a/b 已通过长度差筛选。"""
    prefix = is_prefix_match(a, b)
    dist = levenshtein_distance(a, b)
    homophone, py = sound_similarity(a, b)

    if prefix:
        short, long = (a, b) if len(a) < len(b) else (b, a)
        conf = round(0.75 + 0.25 * (len(short) / len(long)), 2)
        reason = f"前缀关系(「{short}」是「{long}」的前缀)"
        if dist <= 2:
            reason += f",Levenshtein 距离 {dist}"
        return conf, reason

    if homophone and dist <= 2:
        return 0.95 if dist <= 1 else 0.80, f"同音字({py})"

    if dist == 1:
        return 0.85, "Levenshtein 距离 1"
    if dist == 2:
        return 0.70, "Levenshtein 距离 2"
    return None


def _is_strong(s):
    """强配对:同音 / 前缀 / 影子名(一侧 count≤2 且足够相似 conf≥0.85)。

    影子名须叠加高相似度 —— 否则一个 count=1 的名字会和所有人凑成 dist-2 噪声对。
    """
    return (
        s["reason"].startswith("同音")
        or "前缀" in s["reason"]
        or (min(s["a_count"], s["b_count"]) <= 2 and s["confidence"] >= 0.85)
    )


def find_alias_suggestions(conn, include_noise=False):
    """
    扫 daily_reports DISTINCT agent_name,两两配对给 confidence。
    排除已在 agent_aliases 的配对(已人工确认/拒绝)。confidence 降序。
    默认只回强配对(同音/前缀/影子名);include_noise=True 回全部(开发/进阶 review 用)。
    """
    rows = conn.execute(
        "SELECT agent_name, source, COUNT(*) AS c FROM daily_reports "
        "GROUP BY agent_name, source"
    ).fetchall()

    # name -> {sources:set, count:int}
    info = {}
    for r in rows:
        e = info.setdefault(r["agent_name"], {"sources": set(), "count": 0})
        e["sources"].add(r["source"])
        e["count"] += r["c"]
    names = sorted(info.keys())

    # 已记录配对(无向)→ frozenset 集合,跳过
    existing = set()
    for r in conn.execute(
        "SELECT canonical_name, alias_name FROM agent_aliases"
    ).fetchall():
        existing.add(frozenset((r["canonical_name"], r["alias_name"])))

    out = []
    for i in range(len(names)):
        for j in range(i + 1, len(names)):
            a, b = names[i], names[j]
            if frozenset((a, b)) in existing:
                continue
            # 长度差 > 50% 排除(挡「卡比」vs「卡比卡比卡比」)
            la, lb = len(a), len(b)
            if abs(la - lb) / max(la, lb) > 0.5:
                continue
            res = _confidence_and_reason(a, b)
            if res is None:
                continue
            conf, reason = res
            out.append({
                "name_a": a,
                "name_b": b,
                "confidence": conf,
                "reason": reason,
                "a_in_sources": sorted(info[a]["sources"]),
                "b_in_sources": sorted(info[b]["sources"]),
                "a_count": info[a]["count"],
                "b_count": info[b]["count"],
            })
    out.sort(key=lambda x: (-x["confidence"], x["name_a"], x["name_b"]))
    if not include_noise:
        out = [s for s in out if _is_strong(s)]
    return out


# ---------- CRUD ----------
def list_aliases(conn, page=1, page_size=50):
    page = max(1, int(page))
    page_size = max(1, min(500, int(page_size)))
    total = conn.execute("SELECT COUNT(*) AS c FROM agent_aliases").fetchone()["c"]
    rows = conn.execute(
        "SELECT id, canonical_name, alias_name, source, confidence, decided_by, "
        "decided_at, note FROM agent_aliases ORDER BY canonical_name, alias_name "
        "LIMIT ? OFFSET ?",
        (page_size, (page - 1) * page_size),
    ).fetchall()
    return {
        "items": [dict(r) for r in rows],
        "total": total,
        "page": page,
        "page_size": page_size,
    }


def add_alias(conn, canonical_name, alias_name, source=None, note=None,
              confidence=None, decided_by="manual", decided_at=None):
    """加一笔别名映射。UNIQUE(alias_name, source) 撞 → raise ValueError。"""
    canonical_name = (canonical_name or "").strip()
    alias_name = (alias_name or "").strip()
    if not canonical_name or not alias_name:
        raise ValueError("canonical_name 与 alias_name 必填")
    if canonical_name == alias_name:
        raise ValueError("canonical_name 与 alias_name 不可相同")
    src = source or None
    if src not in (None, "yueda", "remote"):
        raise ValueError("source 只能是 yueda / remote / 空")
    decided_at = decided_at or datetime.datetime.now().isoformat()
    try:
        cur = conn.execute(
            "INSERT INTO agent_aliases "
            "(canonical_name, alias_name, source, confidence, decided_by, decided_at, note) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            (canonical_name, alias_name, src, confidence, decided_by, decided_at, note),
        )
        conn.commit()
    except Exception as e:
        # sqlite3.IntegrityError 等
        if "UNIQUE" in str(e):
            raise ValueError(f"别名「{alias_name}」(source={src})已存在映射") from e
        raise
    return {
        "id": cur.lastrowid,
        "canonical_name": canonical_name,
        "alias_name": alias_name,
        "source": src,
        "confidence": confidence,
        "decided_by": decided_by,
        "decided_at": decided_at,
        "note": note,
    }


def remove_alias(conn, alias_id):
    cur = conn.execute("DELETE FROM agent_aliases WHERE id = ?", (alias_id,))
    conn.commit()
    return cur.rowcount > 0
