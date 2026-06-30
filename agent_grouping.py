#!/usr/bin/env python3
"""
agent_grouping.py — Phase 17 段 1

接线量分组字典(硬编码 28 人名单)。

为什么 hardcode:daily_reports.source 只有 'yueda' / 'remote' 两值,
无法据资料推出 6 群(1-1 早/中/夜 全是 yueda;1-1 远程 + 1-2 远程 全是 remote)。
因此分群必须靠 canonical_name → 组别 的人工字典。

镜像 agent_aliases.py 的「纯模块」模式:不碰 DB、不碰 Flask,只做名单映射,
在 build_alias_map 归一后的 canonical_name 上套用。

诚实声明:3 个组长(贝果 / 沙西米 / 轩轩)不提交日报,因此不会出现在
daily_reports → 任何 tab(含综合)都看不到他们,这是资料层的诚实,不是 bug。

新人「尖尖」7 月起加入 1-1 夜班,已预留在字典内;资料一进 daily_reports
就会自动在 11_night 群显示,无需改码。
"""

# ──────────────────────────────────────────────────────────────────────
# 6 群定义(key → 中文说明)。'all' = 综合,不过滤。
# ──────────────────────────────────────────────────────────────────────
GROUP_KEYS = {
    "all":        {"description": "综合(全部接线客服)"},
    "11_morning": {"description": "1-1 早班"},
    "11_mid":     {"description": "1-1 中班"},
    "11_night":   {"description": "1-1 夜班"},
    "11_remote":  {"description": "1-1 远程"},
    "12_remote":  {"description": "1-2 远程"},
}

# 合法的「可过滤」群(不含 all)。供 endpoint 校验用。
FILTERABLE_GROUPS = [k for k in GROUP_KEYS if k != "all"]


# ──────────────────────────────────────────────────────────────────────
# 28 人名单:canonical_name → {dept, location, shift, role, group}
#   dept     '1-1' / '1-2'
#   location 'onsite'(悦达驻场) / 'remote'(远程)
#   shift    'morning' / 'mid' / 'night' / 'remote'
#   role     'leader'(组长,不提交日报) / 'member'
#   group    对应 GROUP_KEYS 的过滤键
# ──────────────────────────────────────────────────────────────────────
ROSTER = {
    # ── 1-1 早班(6 人,onsite/悦达;组长 贝果 不提交日报)──
    "贝果":   {"dept": "1-1", "location": "onsite", "shift": "morning", "role": "leader", "group": "11_morning"},
    "卡比":   {"dept": "1-1", "location": "onsite", "shift": "morning", "role": "member", "group": "11_morning"},
    "艾娃":   {"dept": "1-1", "location": "onsite", "shift": "morning", "role": "member", "group": "11_morning"},
    "路奇":   {"dept": "1-1", "location": "onsite", "shift": "morning", "role": "member", "group": "11_morning"},
    "艾瑞克": {"dept": "1-1", "location": "onsite", "shift": "morning", "role": "member", "group": "11_morning"},
    "果冻":   {"dept": "1-1", "location": "onsite", "shift": "morning", "role": "member", "group": "11_morning"},

    # ── 1-1 中班(6 人,onsite/悦达;组长 沙西米 不提交日报)──
    "沙西米": {"dept": "1-1", "location": "onsite", "shift": "mid", "role": "leader", "group": "11_mid"},
    "卡姆利": {"dept": "1-1", "location": "onsite", "shift": "mid", "role": "member", "group": "11_mid"},
    "小玥":   {"dept": "1-1", "location": "onsite", "shift": "mid", "role": "member", "group": "11_mid"},
    "鱼丸":   {"dept": "1-1", "location": "onsite", "shift": "mid", "role": "member", "group": "11_mid"},
    "翅膀":   {"dept": "1-1", "location": "onsite", "shift": "mid", "role": "member", "group": "11_mid"},
    "霄霄":   {"dept": "1-1", "location": "onsite", "shift": "mid", "role": "member", "group": "11_mid"},

    # ── 1-1 夜班(7 人 + 尖尖 7月起预留,onsite/悦达;组长 轩轩 不提交日报)──
    "轩轩":   {"dept": "1-1", "location": "onsite", "shift": "night", "role": "leader", "group": "11_night"},
    "大雄":   {"dept": "1-1", "location": "onsite", "shift": "night", "role": "member", "group": "11_night"},
    "咖啡":   {"dept": "1-1", "location": "onsite", "shift": "night", "role": "member", "group": "11_night"},
    "小江":   {"dept": "1-1", "location": "onsite", "shift": "night", "role": "member", "group": "11_night"},
    "九节狼": {"dept": "1-1", "location": "onsite", "shift": "night", "role": "member", "group": "11_night"},
    "小邱":   {"dept": "1-1", "location": "onsite", "shift": "night", "role": "member", "group": "11_night"},
    "当肯":   {"dept": "1-1", "location": "onsite", "shift": "night", "role": "member", "group": "11_night"},
    "尖尖":   {"dept": "1-1", "location": "onsite", "shift": "night", "role": "member", "group": "11_night"},  # 7月起加入,资料一到自动显示

    # ── 1-1 远程(5 人,remote)──
    "宋江":   {"dept": "1-1", "location": "remote", "shift": "remote", "role": "member", "group": "11_remote"},
    "鲍博":   {"dept": "1-1", "location": "remote", "shift": "remote", "role": "member", "group": "11_remote"},
    "华乙":   {"dept": "1-1", "location": "remote", "shift": "remote", "role": "member", "group": "11_remote"},
    "尹东":   {"dept": "1-1", "location": "remote", "shift": "remote", "role": "member", "group": "11_remote"},
    "李中意": {"dept": "1-1", "location": "remote", "shift": "remote", "role": "member", "group": "11_remote"},

    # ── 1-2 远程(4 人,remote)──
    "江初雨": {"dept": "1-2", "location": "remote", "shift": "remote", "role": "member", "group": "12_remote"},
    "刘玉虹": {"dept": "1-2", "location": "remote", "shift": "remote", "role": "member", "group": "12_remote"},
    "谢辰":   {"dept": "1-2", "location": "remote", "shift": "remote", "role": "member", "group": "12_remote"},
    "西南":   {"dept": "1-2", "location": "remote", "shift": "remote", "role": "member", "group": "12_remote"},
}

# 每群 canonical_name 集合(预算一次,过滤用)。
GROUP_MEMBERS = {g: {name for name, meta in ROSTER.items() if meta["group"] == g}
                 for g in FILTERABLE_GROUPS}


def group_of(canonical_name):
    """回传 canonical_name 所属群 key,名单外的人回 None(例:谢飞飞、新临时人）。"""
    meta = ROSTER.get(canonical_name)
    return meta["group"] if meta else None


def filter_agents_by_group(agents_table, group_filter):
    """从 agents_table 取出属于 group_filter 的列。

    group_filter='all' 不过滤(原样返回);名单外的人(不在 ROSTER)在任何
    非 all 群都不会出现 —— 综合 tab 才看得到他们。
    """
    if group_filter == "all":
        return agents_table
    members = GROUP_MEMBERS.get(group_filter, set())
    return [a for a in agents_table if a.get("canonical_name") in members]
