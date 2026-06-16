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
    expected = {"refunds", "uploads", "daily_reports"}
    if not expected.issubset(set(tables)):
        raise SystemExit(f"ERROR: missing tables, expected {expected}, got {tables}")
    print("OK: refunds + uploads + daily_reports present.")
