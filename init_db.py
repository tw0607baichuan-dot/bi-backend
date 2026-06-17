#!/usr/bin/env python3
"""
SQLite 资料库初始化脚本

Phase 2.2 — refunds + uploads(退费子系统)
Phase 3.1 — daily_reports(接线量子页:悦达 3 班 + 远程日报)

建立 /var/data/refunds/db.sqlite。可重复执行(IF NOT EXISTS),不会清掉既有资料。
"""
import os
import sqlite3

DB_PATH = "/var/data/refunds/db.sqlite"

SCHEMA = """
CREATE TABLE IF NOT EXISTS refunds (
    id INTEGER PRIMARY KEY AUTOINCREMENT,

    -- Excel 原始 7 栏(完全对照)
    date_iso TEXT NOT NULL,           -- ISO 格式 YYYY-MM-DD
    merchant_tg TEXT NOT NULL,        -- 悦达 / 悦达2 等 TG 号
    merchant_order_no TEXT NOT NULL,  -- dbs2_xxxxx
    platform_order_no TEXT NOT NULL,  -- WNSYxxxxx
    amount INTEGER NOT NULL,          -- 单笔金额(元)
    payment_type TEXT NOT NULL,       -- 支付宝 / 微信
    platform_id TEXT NOT NULL,        -- nineone / tiktok 等

    -- 上传 metadata
    upload_id INTEGER NOT NULL,       -- FK 到 uploads.id

    -- 索引
    UNIQUE(merchant_order_no)         -- 同一笔退费只能存在一次(防重复上传)
);

CREATE INDEX IF NOT EXISTS idx_refunds_date ON refunds(date_iso);
CREATE INDEX IF NOT EXISTS idx_refunds_platform ON refunds(platform_id);
CREATE INDEX IF NOT EXISTS idx_refunds_upload ON refunds(upload_id);

CREATE TABLE IF NOT EXISTS uploads (
    id INTEGER PRIMARY KEY AUTOINCREMENT,

    -- 档案 metadata
    original_filename TEXT NOT NULL,   -- 上传时的原始档名
    stored_filename TEXT NOT NULL,     -- 储存后的档名(含 timestamp)
    file_size_bytes INTEGER NOT NULL,
    file_md5 TEXT NOT NULL,            -- 用来侦测重复上传

    -- 业务 metadata(从档名 + 内容解析)
    product_line TEXT NOT NULL,        -- 实际意义是「公司」(目前全部是「悦达」),保留字段名是历史原因
    week_start TEXT NOT NULL,          -- 该档案涵盖的起始日
    week_end TEXT NOT NULL,            -- 该档案涵盖的结束日
    rows_imported INTEGER NOT NULL,    -- 解析进 SQLite 几笔
    rows_skipped INTEGER NOT NULL,     -- 因 UNIQUE 重复跳过几笔

    -- 上传者 + 时间
    uploaded_by TEXT NOT NULL,         -- 用户名
    uploaded_role TEXT NOT NULL,       -- 当时角色(superadmin / leader 等)
    uploaded_at TEXT NOT NULL,         -- ISO timestamp

    -- 状态
    status TEXT NOT NULL DEFAULT 'active'  -- active / superseded / deleted
);

CREATE INDEX IF NOT EXISTS idx_uploads_product_week ON uploads(product_line, week_start, week_end);
CREATE INDEX IF NOT EXISTS idx_uploads_status ON uploads(status);

-- Phase 3.1 — 接线量子页:坐席日报(悦达 3 班 + 远程日报)
CREATE TABLE IF NOT EXISTS daily_reports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,

    -- 日期 + 提交人(核心键)
    date_iso TEXT NOT NULL,              -- 修正后的日期(可能 = original)
    original_date_iso TEXT,              -- 原始填写日期(被修正才有值,否则 NULL)
    correction_note TEXT,                -- 修正说明,例 "原始抢贴 2026-06-02 → 修正为 2026-06-03"

    agent_name TEXT NOT NULL,

    -- 业务指标(全部可 NULL,因为原始数据可能脏)
    intake INTEGER,
    response_time_sec INTEGER,
    quality_score REAL,
    escalation_count INTEGER,

    -- 班次保留(未来扩展)
    shift TEXT,                          -- 早 / 中 / 晚 / 白 / 夜

    -- 来源
    source TEXT NOT NULL,                -- 'yueda' / 'remote'

    -- 元数据
    submit_time TEXT,                    -- 提交时间 HH:MM
    record_time TEXT,                    -- 收錄時間(仅远程有,用来推断修正日期)
    raw_data TEXT,                       -- 原始 row JSON
    synced_at TEXT NOT NULL,

    UNIQUE(source, date_iso, agent_name) -- 修正后零碰撞
);

CREATE INDEX IF NOT EXISTS idx_dr_date ON daily_reports(date_iso);
CREATE INDEX IF NOT EXISTS idx_dr_agent ON daily_reports(agent_name);
CREATE INDEX IF NOT EXISTS idx_dr_source ON daily_reports(source);
CREATE INDEX IF NOT EXISTS idx_dr_corrected ON daily_reports(correction_note) WHERE correction_note IS NOT NULL;

-- Phase 3.2 — 人名合并(别名归一):daily_reports 原始名不动,查询时 COALESCE 到 canonical
CREATE TABLE IF NOT EXISTS agent_aliases (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    canonical_name TEXT NOT NULL,         -- 主名(查询时归一到这个)
    alias_name TEXT NOT NULL,             -- 别名(daily_reports 里出现的原名)
    source TEXT,                          -- 'yueda' / 'remote' / NULL(跨源都映射)
    confidence REAL,                      -- 0.0-1.0(fuzzy 推断给的)
    decided_by TEXT NOT NULL,             -- 'fuzzy_auto' / 'manual_baichuan' / 'manual_other'
    decided_at TEXT NOT NULL,
    note TEXT,                            -- 决策说明(例:'同音字归一')
    UNIQUE(alias_name, source)            -- 一个别名只能对应一个主名
);

CREATE INDEX IF NOT EXISTS idx_aa_canonical ON agent_aliases(canonical_name);
CREATE INDEX IF NOT EXISTS idx_aa_alias ON agent_aliases(alias_name);

-- Phase 4.1 — 质检子页:错误案例总表 + 客服当日汇总 + uploads metadata
-- 镜像 refunds/uploads 模式(supersede + UNIQUE 防呆 + parser 独立档案)。
-- agent_name 原值保留(查询时再 JOIN agent_aliases 归一,不在写入时归一以保溯源)。

-- 1. 错误案例总表(每笔 = 一个具体扣分案例)
CREATE TABLE IF NOT EXISTS quality_inspections (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    upload_id INTEGER NOT NULL,            -- FK to quality_uploads

    -- 审核维度
    inspect_date TEXT NOT NULL,            -- 「审核日期」三层降级解出
    dept TEXT NOT NULL,                    -- 'dx' / 'df'(本期写死 'dx')

    -- 案例核心
    case_no INTEGER,                       -- 序号(Excel 第一栏)
    shift TEXT,                            -- 班别(早/中/晚/夜)
    agent_name TEXT NOT NULL,              -- 客服中文名(原值保留)
    agent_account TEXT,                    -- 客服账号(英文 id 例 nine9)
    case_time TEXT,                        -- HH:MM(原始,不转 datetime)

    -- 产品上下文
    app_name TEXT,                         -- 完整 App 名(例「小蓝(DX-030)」)
    app_code TEXT,                         -- 抽出来的产品 ID(例「DX-030」),无则 NULL
    session_id TEXT,                       -- 会话 ID
    user_uid TEXT,                         -- 用户 UID

    -- 扣分
    error_level TEXT NOT NULL,             -- '严重' / '中等' / '轻微'
    deduction REAL NOT NULL,               -- 统一存负数 -3.0 / -1.5 / -0.6

    -- 上下文(给主管 / 组员看)
    error_desc TEXT,                       -- 错误描述
    correct_reply TEXT,                    -- 正确回复方式
    conversation TEXT,                     -- 完整对话上下文(超长文,不截断)

    synced_at TEXT NOT NULL,               -- 写入时间

    UNIQUE(upload_id, case_no)             -- 同上传里序号唯一,避免重复
);

CREATE INDEX IF NOT EXISTS idx_qi_date ON quality_inspections(inspect_date);
CREATE INDEX IF NOT EXISTS idx_qi_agent ON quality_inspections(agent_name);
CREATE INDEX IF NOT EXISTS idx_qi_dept ON quality_inspections(dept);
CREATE INDEX IF NOT EXISTS idx_qi_level ON quality_inspections(error_level);
CREATE INDEX IF NOT EXISTS idx_qi_session ON quality_inspections(session_id);

-- 2. 客服当日汇总
CREATE TABLE IF NOT EXISTS quality_summary (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    upload_id INTEGER NOT NULL,

    inspect_date TEXT NOT NULL,
    dept TEXT NOT NULL,                    -- 'dx' / 'df'

    shift TEXT,                            -- 班别(允许「名单外」)
    agent_name TEXT NOT NULL,              -- 原值保留
    agent_account TEXT,

    total_messages INTEGER,                -- 总讯息数
    severe_count INTEGER DEFAULT 0,        -- 严重次数
    medium_count INTEGER DEFAULT 0,        -- 中等次数
    minor_count INTEGER DEFAULT 0,         -- 轻微次数
    deduction_sum REAL DEFAULT 0,          -- 扣分总和(正数,例 6.0)
    pass_rate REAL,                        -- 合格率(0-1,例 0.9744)
    note TEXT,                             -- 备注(「全合格」或 NULL)

    synced_at TEXT NOT NULL,

    UNIQUE(upload_id, agent_name)          -- 同上传同人唯一
);

CREATE INDEX IF NOT EXISTS idx_qs_date ON quality_summary(inspect_date);
CREATE INDEX IF NOT EXISTS idx_qs_agent ON quality_summary(agent_name);
CREATE INDEX IF NOT EXISTS idx_qs_dept ON quality_summary(dept);

-- 3. Uploads metadata(同 refunds.uploads 模式)
CREATE TABLE IF NOT EXISTS quality_uploads (
    id INTEGER PRIMARY KEY AUTOINCREMENT,

    -- 档案
    original_filename TEXT,
    stored_filename TEXT,
    file_size INTEGER,
    md5 TEXT NOT NULL UNIQUE,              -- 防完全相同档案重复上传

    -- 解析结果
    inspect_date TEXT NOT NULL,            -- 「审核日期」(关键!supersede 基于这字段)
    inspect_date_source TEXT,              -- 'filename' / 'metadata' / 'upload_time'(降级标识)
    dept TEXT NOT NULL,                    -- 'dx' / 'df'

    -- 笔数统计
    inspections_count INTEGER,             -- Sheet 1 入了几笔
    summary_count INTEGER,                 -- Sheet 2 入了几笔

    -- 谁传的 / 何时
    uploaded_by_role TEXT NOT NULL,
    uploaded_by_user TEXT NOT NULL,        -- 英文 username
    uploaded_at TEXT NOT NULL,

    -- 生命周期
    status TEXT NOT NULL DEFAULT 'active', -- 'active' / 'superseded' / 'deleted'
    superseded_by INTEGER,                 -- supersede 时记录被哪个 upload 覆盖
    superseded_at TEXT
);

CREATE INDEX IF NOT EXISTS idx_qu_inspect_date ON quality_uploads(inspect_date);
CREATE INDEX IF NOT EXISTS idx_qu_status ON quality_uploads(status);
CREATE INDEX IF NOT EXISTS idx_qu_dept ON quality_uploads(dept);
"""


def init_db(db_path=DB_PATH):
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    conn = sqlite3.connect(db_path)
    try:
        conn.executescript(SCHEMA)
        conn.commit()
        tables = [
            row[0]
            for row in conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
            ).fetchall()
        ]
        indexes = [
            row[0]
            for row in conn.execute(
                "SELECT name FROM sqlite_master WHERE type='index' "
                "AND name LIKE 'idx_%' ORDER BY name"
            ).fetchall()
        ]
    finally:
        conn.close()
    return tables, indexes


if __name__ == "__main__":
    tables, indexes = init_db()
    print(f"DB initialized at {DB_PATH}")
    print(f"Tables : {tables}")
    print(f"Indexes: {indexes}")
    expected = {
        "refunds", "uploads", "daily_reports", "agent_aliases",
        "quality_inspections", "quality_summary", "quality_uploads",
    }
    if not expected.issubset(set(tables)):
        raise SystemExit(f"ERROR: missing tables, expected {expected}, got {tables}")
    print("OK: refunds + uploads + daily_reports + quality_* present.")
