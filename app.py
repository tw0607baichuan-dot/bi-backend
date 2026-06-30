#!/usr/bin/env python3
"""
BI Dashboard Backend — Phase 2.2

保留 Phase 2.1 的 /api/health,新增:
  POST /api/refunds/upload  — 上传 Excel 退费档(权限验证 + 存原档 + 解析进 SQLite + 旧版归档)
  GET  /api/refunds/list    — 列出所有 status='active' 的上传 metadata

临时验证:用 header X-User-Role(2.3 改真实 session)。
"""
import datetime
import hashlib
import io
import json
import os
import re
import sqlite3

from flask import Flask, jsonify, request

import agent_aliases
import agent_queries
import agents_sync
import excel_parser
import quality_parser
import quality_queries
import refund_queries

app = Flask(__name__)

# ---------- 设定 ----------
DATA_DIR = "/var/data/refunds"
DB_PATH = os.path.join(DATA_DIR, "db.sqlite")
RAW_DIR = os.path.join(DATA_DIR, "raw")
ARCHIVE_DIR = os.path.join(DATA_DIR, "archive")

# Phase 4.1 — 质检子系统(独立 data 目录,同 refunds 模式)
QUALITY_DIR = "/var/data/quality"
QUALITY_RAW_DIR = os.path.join(QUALITY_DIR, "raw")
QUALITY_ARCHIVE_DIR = os.path.join(QUALITY_DIR, "archive")

# Phase 13-2:月度排名「低访问量」门槛。讯息 < 此值自动进 low_volume,不进主榜。
# 完全靠门槛过滤(不维护写死排除清单),新 Excel 帐号按讯息量自动分流,0 维护成本。
QUALITY_RANKING_MIN_MESSAGES_DEFAULT = 50

ALLOWED_ROLES = {"superadmin", "manager", "maintainer", "leader"}


# ---------- 共用 ----------
def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def _safe_component(s):
    """把字串清成可安全用于档名的片段(挡路径穿越 / 非法字元)。"""
    s = str(s).strip()
    s = re.sub(r"[^\w一-鿿.-]", "_", s)  # 保留中英数 / 底线 / 点 / 减号
    return s or "x"


# ---------- Phase 2.1 ----------
@app.route("/api/health")
def health():
    return jsonify({
        "status": "ok",
        "version": "phase-4.2",
        "timestamp": datetime.datetime.now().isoformat(),
    })


# ---------- Phase 2.2 ----------
@app.route("/api/refunds/upload", methods=["POST"])
def upload_refunds():
    # 1) 权限(临时:X-User-Role header)
    role = request.headers.get("X-User-Role")
    if not role:
        return jsonify({"status": "error", "message": "缺 X-User-Role header"}), 401
    if role not in ALLOWED_ROLES:
        return jsonify({
            "status": "error",
            "message": f"角色 {role} 无上传权限",
        }), 403
    uploaded_by = request.headers.get("X-User-Name") or role

    # 2) 取档
    if "file" not in request.files:
        return jsonify({"status": "error", "message": "未带 file"}), 400
    f = request.files["file"]
    original_filename = f.filename or ""
    if not original_filename:
        return jsonify({"status": "error", "message": "档名为空"}), 400

    file_bytes = f.read()
    if not file_bytes:
        return jsonify({"status": "error", "message": "档案为空"}), 400
    file_md5 = hashlib.md5(file_bytes).hexdigest()
    file_size = len(file_bytes)

    # 3) md5 重复上传检查(在写档前,避免产生孤儿档)
    conn = get_conn()
    try:
        dup = conn.execute(
            "SELECT id FROM uploads WHERE file_md5 = ? AND status != 'deleted'",
            (file_md5,),
        ).fetchone()
        if dup:
            return jsonify({
                "status": "error",
                "message": f"档案已上传过(md5 重复,upload_id={dup['id']})",
            }), 409

        # 4) 解析 Excel(先解析才知道年份 / 产品线 / 周次)
        try:
            rows = excel_parser.parse_excel(io.BytesIO(file_bytes))
            year = excel_parser.first_date_year(rows)
            meta = excel_parser.parse_filename(original_filename, default_year=year)
        except excel_parser.FilenameParseError as e:
            return jsonify({"status": "error", "message": f"档名解析失败:{e}"}), 400
        except excel_parser.MissingColumnError as e:
            return jsonify({"status": "error", "message": f"栏位缺失:{e}"}), 400
        except excel_parser.FileCorruptError as e:
            return jsonify({"status": "error", "message": f"档案损坏:{e}"}), 400
        except excel_parser.ExcelParseError as e:
            return jsonify({"status": "error", "message": f"解析失败:{e}"}), 400

        product_line = meta["product_line"]
        week_start = meta["week_start"]
        week_end = meta["week_end"]

        # 5) 储存原档(档名由 server 生成,挡路径穿越)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        stored_filename = (
            f"{_safe_component(product_line)}_{_safe_component(week_start)}_"
            f"{_safe_component(week_end)}_{timestamp}.xlsx"
        )
        os.makedirs(RAW_DIR, exist_ok=True)
        os.makedirs(ARCHIVE_DIR, exist_ok=True)
        stored_path = os.path.join(RAW_DIR, stored_filename)
        with open(stored_path, "wb") as out:
            out.write(file_bytes)

        uploaded_at = datetime.datetime.now().isoformat()

        # 6) 先建 upload 列(拿 upload_id)
        cur = conn.execute(
            """
            INSERT INTO uploads (
                original_filename, stored_filename, file_size_bytes, file_md5,
                product_line, week_start, week_end, rows_imported, rows_skipped,
                uploaded_by, uploaded_role, uploaded_at, status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, 0, 0, ?, ?, ?, 'active')
            """,
            (
                original_filename, stored_filename, file_size, file_md5,
                product_line, week_start, week_end,
                uploaded_by, role, uploaded_at,
            ),
        )
        upload_id = cur.lastrowid

        # 7) INSERT OR IGNORE refunds(重复 merchant_order_no 自动跳过)
        before = conn.execute("SELECT COUNT(*) AS c FROM refunds").fetchone()["c"]
        conn.executemany(
            """
            INSERT OR IGNORE INTO refunds (
                date_iso, merchant_tg, merchant_order_no, platform_order_no,
                amount, payment_type, platform_id, upload_id
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    r["date_iso"], r["merchant_tg"], r["merchant_order_no"],
                    r["platform_order_no"], r["amount"], r["payment_type"],
                    r["platform_id"], upload_id,
                )
                for r in rows
            ],
        )
        after = conn.execute("SELECT COUNT(*) AS c FROM refunds").fetchone()["c"]
        rows_imported = after - before
        rows_skipped = len(rows) - rows_imported

        conn.execute(
            "UPDATE uploads SET rows_imported = ?, rows_skipped = ? WHERE id = ?",
            (rows_imported, rows_skipped, upload_id),
        )

        # 8) 旧版归档:同产品线 + 同周次的旧 active → superseded,原档移 archive/
        olds = conn.execute(
            """
            SELECT id, stored_filename FROM uploads
            WHERE product_line = ? AND week_start = ? AND week_end = ?
              AND status = 'active' AND id != ?
            """,
            (product_line, week_start, week_end, upload_id),
        ).fetchall()
        superseded_ids = []
        for old in olds:
            conn.execute(
                "UPDATE uploads SET status = 'superseded' WHERE id = ?", (old["id"],)
            )
            superseded_ids.append(old["id"])
            old_path = os.path.join(RAW_DIR, old["stored_filename"])
            if os.path.exists(old_path):
                try:
                    os.replace(old_path, os.path.join(ARCHIVE_DIR, old["stored_filename"]))
                except OSError:
                    pass  # 归档移动失败不影响主流程(原档仍在 raw/)

        conn.commit()
    finally:
        conn.close()

    return jsonify({
        "status": "ok",
        "upload_id": upload_id,
        "rows_imported": rows_imported,
        "rows_skipped": rows_skipped,
        "product_line": product_line,
        "week_start": week_start,
        "week_end": week_end,
        "superseded": superseded_ids,
        "message": f"汇入 {rows_imported} 笔,跳过 {rows_skipped} 笔",
    }), 200


@app.route("/api/refunds/list")
def list_refunds():
    conn = get_conn()
    try:
        rows = conn.execute(
            """
            SELECT id, original_filename, stored_filename, file_size_bytes, file_md5,
                   product_line, week_start, week_end, rows_imported, rows_skipped,
                   uploaded_by, uploaded_role, uploaded_at, status
            FROM uploads
            WHERE status = 'active'
            ORDER BY uploaded_at DESC
            """
        ).fetchall()
    finally:
        conn.close()
    return jsonify([dict(r) for r in rows])


def _has_active_data(conn):
    """有 active upload 且 refunds 非空才算有资料。"""
    active = conn.execute(
        "SELECT COUNT(*) AS c FROM uploads WHERE status = 'active'"
    ).fetchone()["c"]
    if active == 0:
        return False
    # 只算 active upload 的 refunds(排除被 supersede 的旧版本明细)
    refunds = conn.execute(
        "SELECT COUNT(*) AS c FROM refunds "
        "WHERE upload_id IN (SELECT id FROM uploads WHERE status = 'active')"
    ).fetchone()["c"]
    return refunds > 0


@app.route("/api/refunds/data")
def refunds_data():
    range_key = request.args.get("range", "week")
    start_date = request.args.get("start_date")
    end_date = request.args.get("end_date")
    compare_key = request.args.get("compare", "none")

    # 时间范围解析(参数错 → 400)
    try:
        start_iso, end_iso, range_label = refund_queries.resolve_time_range(
            range_key, start_date, end_date
        )
        compare_range = refund_queries.resolve_compare_range(
            compare_key, start_iso, end_iso
        )
    except ValueError as e:
        return jsonify({"status": "error", "message": str(e)}), 400

    range_block = {"label": range_label, "start": start_iso, "end": end_iso}

    conn = get_conn()
    try:
        # DB 空 → 200 + 空结构(让前端显示「请上传 Excel」)
        if not _has_active_data(conn):
            return jsonify({
                "status": "ok",
                "has_data": False,
                "range": range_block,
                "current": refund_queries.empty_period(),
                "compare": None,
                "insights": [],
            })

        current = refund_queries.build_period(conn, start_iso, end_iso)

        compare = None
        if compare_range is not None:
            c_start, c_end, c_label = compare_range
            compare = refund_queries.build_period(conn, c_start, c_end)
            compare["label"] = c_label
            compare["start"] = c_start
            compare["end"] = c_end

        insights = refund_queries.compute_insights(current, compare)
    finally:
        conn.close()

    return jsonify({
        "status": "ok",
        "has_data": True,
        "range": range_block,
        "current": current,
        "compare": compare,
        "insights": insights,
    })


@app.route("/api/refunds/health-data")
def refunds_health_data():
    """前端预热用:回报 DB 是否有资料 + 资料涵盖区间。"""
    conn = get_conn()
    try:
        has_data = _has_active_data(conn)
        # total_refunds 与涵盖区间都只看 active upload 的明细(排除 superseded,与 KPI 端点一致)
        active_subq = "WHERE upload_id IN (SELECT id FROM uploads WHERE status = 'active')"
        total = conn.execute(
            f"SELECT COUNT(*) AS c FROM refunds {active_subq}"
        ).fetchone()["c"]
        active = conn.execute(
            "SELECT COUNT(*) AS c FROM uploads WHERE status = 'active'"
        ).fetchone()["c"]
        span = conn.execute(
            f"SELECT MIN(date_iso) AS mn, MAX(date_iso) AS mx FROM refunds {active_subq}"
        ).fetchone()
    finally:
        conn.close()
    return jsonify({
        "status": "ok",
        "has_data": has_data,
        "total_refunds": total,
        "active_uploads": active,
        "earliest": span["mn"],
        "latest": span["mx"],
    })


# ---------- Phase 3.1 — 接线量子页:坐席日报 ----------
@app.route("/api/agents/sync", methods=["POST"])
def agents_sync_endpoint():
    """手动触发同步(权限 leader+);cron 也打这个。失败 raise → 5xx。"""
    role = request.headers.get("X-User-Role")
    if not role:
        return jsonify({"status": "error", "message": "缺 X-User-Role header"}), 401
    if role not in ALLOWED_ROLES:
        return jsonify({"status": "error", "message": f"角色 {role} 无同步权限"}), 403

    conn = get_conn()
    try:
        result = agents_sync.sync_all(conn)
    finally:
        conn.close()

    return jsonify({
        "status": "ok",
        "yueda_imported": result["yueda_imported"],
        "yueda_skipped": result["yueda_skipped"],
        "remote_imported": result["remote_imported"],
        "remote_skipped": result["remote_skipped"],
        "remote_corrected": result["remote_corrected"],
        "corrections": result["corrections"],
        "warnings": result["warnings"],
        "total": result["total"],
        "synced_at": result["synced_at"],
        "errors": [],
    }), 200


@app.route("/api/agents/health-data")
def agents_health_data():
    """前端预热用:各 source 笔数 / 日期范围 / 远程修正数 / 最近同步 / agent 数。"""
    conn = get_conn()
    try:
        sources = {}
        for src in ("yueda", "remote"):
            row = conn.execute(
                """
                SELECT COUNT(*) AS count,
                       COUNT(DISTINCT agent_name) AS agents,
                       MIN(date_iso) AS earliest,
                       MAX(date_iso) AS latest,
                       COUNT(correction_note) AS corrected
                FROM daily_reports WHERE source = ?
                """,
                (src,),
            ).fetchone()
            entry = {
                "count": row["count"],
                "agents": row["agents"],
                "earliest": row["earliest"],
                "latest": row["latest"],
            }
            if src == "remote":
                entry["corrected"] = row["corrected"]
            sources[src] = entry

        agg = conn.execute(
            "SELECT MAX(synced_at) AS synced_at, COUNT(DISTINCT agent_name) AS agents, "
            "COUNT(*) AS total FROM daily_reports"
        ).fetchone()
    finally:
        conn.close()

    return jsonify({
        "status": "ok",
        "sources": sources,
        "total": agg["total"],
        "distinct_agents": agg["agents"],
        "synced_at": agg["synced_at"],
    })


# ---------- Phase 3.2 — 接线量子页:查询 API + 人名合并 ----------
def _check_role():
    """leader+ 权限检查。通过回 (role, name);否则回 (error_json, status)。"""
    role = request.headers.get("X-User-Role")
    if not role:
        return None, (jsonify({"status": "error", "message": "缺 X-User-Role header"}), 401)
    if role not in ALLOWED_ROLES:
        return None, (jsonify({"status": "error", "message": f"角色 {role} 无权限"}), 403)
    name = request.headers.get("X-User-Name") or role
    return (role, name), None


@app.route("/api/agents/data")
def agents_data():
    """主查询:KPI 5 卡 + 趋势 + 客服明细 + 热力图 + 异常名单 + 守门说明。"""
    range_key = request.args.get("range", "week")
    start_date = request.args.get("start_date")
    end_date = request.args.get("end_date")
    compare_key = request.args.get("compare", "none")
    source_filter = request.args.get("source", "all")
    if source_filter not in ("all", "yueda", "remote"):
        return jsonify({"status": "error", "message": f"未知 source: {source_filter}"}), 400
    try:
        absence_days = int(request.args.get("anomaly_days", agent_queries.DEFAULT_ABSENCE_DAYS))
    except ValueError:
        return jsonify({"status": "error", "message": "anomaly_days 须为整数"}), 400

    try:
        start_iso, end_iso, range_label = agent_queries.resolve_time_range(
            range_key, start_date, end_date
        )
        compare_range = agent_queries.resolve_compare_range(compare_key, start_iso, end_iso)
    except ValueError as e:
        return jsonify({"status": "error", "message": str(e)}), 400

    conn = get_conn()
    try:
        amap = agent_aliases.build_alias_map(conn)
        current = agent_queries.build_period(
            conn, start_iso, end_iso, source_filter, amap, full=True, absence_days=absence_days
        )
        compare = None
        if compare_range is not None:
            c_start, c_end, c_label = compare_range
            compare = agent_queries.build_period(
                conn, c_start, c_end, source_filter, amap, full=False
            )
            compare["label"] = c_label
            compare["start"] = c_start
            compare["end"] = c_end
    finally:
        conn.close()

    return jsonify({
        "status": "ok",
        "range": {"label": range_label, "start": start_iso, "end": end_iso},
        "source_filter": source_filter,
        "current": current,
        "compare": compare,
    })


@app.route("/api/agents/alias-suggestions")
def agents_alias_suggestions():
    """人名 fuzzy 配对建议(已排除 agent_aliases 里已决策的配对)。

    默认只回强配对(同音/前缀/影子名);?include_noise=true 回全部。
    """
    include_noise = request.args.get("include_noise", "").lower() in ("1", "true", "yes")
    conn = get_conn()
    try:
        suggestions = agent_aliases.find_alias_suggestions(conn, include_noise=include_noise)
    finally:
        conn.close()
    return jsonify({
        "status": "ok",
        "count": len(suggestions),
        "include_noise": include_noise,
        "suggestions": suggestions,
    })


@app.route("/api/agents/aliases", methods=["GET"])
def agents_aliases_list():
    """列出已确认的 alias 配对(分页)。"""
    try:
        page = int(request.args.get("page", 1))
        page_size = int(request.args.get("page_size", 50))
    except ValueError:
        return jsonify({"status": "error", "message": "page / page_size 须为整数"}), 400
    conn = get_conn()
    try:
        result = agent_aliases.list_aliases(conn, page, page_size)
    finally:
        conn.close()
    result["status"] = "ok"
    return jsonify(result)


@app.route("/api/agents/aliases", methods=["POST"])
def agents_aliases_add():
    """加一笔 alias(权限 leader+)。"""
    auth, err = _check_role()
    if err:
        return err
    role, name = auth
    body = request.get_json(silent=True) or {}
    decided_by = body.get("decided_by") or f"manual_{name}"
    try:
        conn = get_conn()
        try:
            row = agent_aliases.add_alias(
                conn,
                canonical_name=body.get("canonical_name"),
                alias_name=body.get("alias_name"),
                source=body.get("source"),
                note=body.get("note"),
                confidence=body.get("confidence"),
                decided_by=decided_by,
            )
        finally:
            conn.close()
    except ValueError as e:
        return jsonify({"status": "error", "message": str(e)}), 400
    return jsonify({"status": "ok", "alias": row}), 200


@app.route("/api/agents/aliases/<int:alias_id>", methods=["DELETE"])
def agents_aliases_remove(alias_id):
    """撤回 alias(权限 leader+)。"""
    auth, err = _check_role()
    if err:
        return err
    conn = get_conn()
    try:
        removed = agent_aliases.remove_alias(conn, alias_id)
    finally:
        conn.close()
    if not removed:
        return jsonify({"status": "error", "message": f"alias id={alias_id} 不存在"}), 404
    return jsonify({"status": "ok", "removed_id": alias_id}), 200


# ════════════════════════════════════════════════════════════════
#  Phase 4.1 — 质检子页:上传 + 解析 + 健康检查 + uploads 清单
#  镜像 refunds 上传模式(supersede + UNIQUE 防呆 + 独立 parser)。
#  agent_name 原值保留,查询归一留 4.2。
# ════════════════════════════════════════════════════════════════
@app.route("/api/quality/upload", methods=["POST"])
def upload_quality():
    # 1) 权限(leader+,同退费)
    role = request.headers.get("X-User-Role")
    if not role:
        return jsonify({"status": "error", "message": "缺 X-User-Role header"}), 401
    if role not in ALLOWED_ROLES:
        return jsonify({"status": "error", "message": f"角色 {role} 无上传权限"}), 403
    uploaded_by = request.headers.get("X-User-Name") or role

    # 2) 取档 → bytes → md5 + size
    if "file" not in request.files:
        return jsonify({"status": "error", "message": "未带 file"}), 400
    f = request.files["file"]
    original_filename = f.filename or ""
    if not original_filename:
        return jsonify({"status": "error", "message": "档名为空"}), 400
    file_bytes = f.read()
    if not file_bytes:
        return jsonify({"status": "error", "message": "档案为空"}), 400
    file_md5 = hashlib.md5(file_bytes).hexdigest()
    file_size = len(file_bytes)

    now_iso = datetime.datetime.now().isoformat()
    dept = "dx"   # 本期写死(1-1 部 DX 系)

    conn = get_conn()
    try:
        # 3) md5 重复上传检查(写档前)
        dup = conn.execute(
            "SELECT id FROM quality_uploads WHERE md5 = ? AND status != 'deleted'",
            (file_md5,),
        ).fetchone()
        if dup:
            return jsonify({
                "status": "error",
                "message": f"档案已上传过(md5 重复,upload_id={dup['id']})",
            }), 409

        # 4) 「审核日期」三层降级
        inspect_date, date_source = quality_parser.resolve_inspect_date(
            original_filename, file_bytes, now_iso
        )

        # 5) 解析 Excel(3 sheet)
        try:
            parsed = quality_parser.parse_quality_excel(file_bytes)
        except quality_parser.QualityParseError as e:
            return jsonify({"status": "error", "message": f"解析失败:{e}"}), 400
        except Exception as e:
            return jsonify({"status": "error", "message": f"档案损坏 / 无法解析:{e}"}), 400

        s1 = parsed["sheet1_rows"]
        s2 = parsed["sheet2_rows"]

        # 8) 存原档(server 生成档名,防路径穿越)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        stored_filename = (
            f"{_safe_component(inspect_date)}_{_safe_component(dept)}_{timestamp}.xlsx"
        )
        os.makedirs(QUALITY_RAW_DIR, exist_ok=True)
        os.makedirs(QUALITY_ARCHIVE_DIR, exist_ok=True)
        stored_path = os.path.join(QUALITY_RAW_DIR, stored_filename)
        with open(stored_path, "wb") as out:
            out.write(file_bytes)

        # 10) INSERT quality_uploads 取 upload_id
        cur = conn.execute(
            """
            INSERT INTO quality_uploads (
                original_filename, stored_filename, file_size, md5,
                inspect_date, inspect_date_source, dept,
                inspections_count, summary_count,
                uploaded_by_role, uploaded_by_user, uploaded_at, status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, 0, 0, ?, ?, ?, 'active')
            """,
            (
                original_filename, stored_filename, file_size, file_md5,
                inspect_date, date_source, dept,
                role, uploaded_by, now_iso,
            ),
        )
        upload_id = cur.lastrowid

        # 11) INSERT quality_inspections(UNIQUE(upload_id, case_no) 防呆)
        conn.executemany(
            """
            INSERT OR IGNORE INTO quality_inspections (
                upload_id, inspect_date, dept, case_no, shift,
                agent_name, agent_account, case_time, app_name, app_code,
                session_id, user_uid, error_level, deduction,
                error_desc, correct_reply, conversation, synced_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    upload_id, inspect_date, dept, r["case_no"], r["shift"],
                    r["agent_name"], r["agent_account"], r["case_time"],
                    r["app_name"], r["app_code"], r["session_id"], r["user_uid"],
                    r["error_level"], r["deduction"],
                    r["error_desc"], r["correct_reply"], r["conversation"], now_iso,
                )
                for r in s1
            ],
        )
        inspections_count = conn.execute(
            "SELECT COUNT(*) AS c FROM quality_inspections WHERE upload_id = ?", (upload_id,)
        ).fetchone()["c"]

        # 12) INSERT quality_summary(UNIQUE(upload_id, agent_name) 防呆)
        conn.executemany(
            """
            INSERT OR IGNORE INTO quality_summary (
                upload_id, inspect_date, dept, shift, agent_name, agent_account,
                total_messages, severe_count, medium_count, minor_count,
                deduction_sum, pass_rate, note, synced_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    upload_id, inspect_date, dept, r["shift"], r["agent_name"],
                    r["agent_account"], r["total_messages"], r["severe_count"],
                    r["medium_count"], r["minor_count"], r["deduction_sum"],
                    r["pass_rate"], r["note"], now_iso,
                )
                for r in s2
            ],
        )
        summary_count = conn.execute(
            "SELECT COUNT(*) AS c FROM quality_summary WHERE upload_id = ?", (upload_id,)
        ).fetchone()["c"]

        conn.execute(
            "UPDATE quality_uploads SET inspections_count = ?, summary_count = ? WHERE id = ?",
            (inspections_count, summary_count, upload_id),
        )

        # 9) supersede:同 inspect_date + dept 的旧 active → superseded,原档归档
        olds = conn.execute(
            """
            SELECT id, stored_filename FROM quality_uploads
            WHERE inspect_date = ? AND dept = ? AND status = 'active' AND id != ?
            """,
            (inspect_date, dept, upload_id),
        ).fetchall()
        superseded_old_id = None
        for old in olds:
            conn.execute(
                "UPDATE quality_uploads SET status = 'superseded', "
                "superseded_by = ?, superseded_at = ? WHERE id = ?",
                (upload_id, now_iso, old["id"]),
            )
            superseded_old_id = old["id"]   # 通常只有一笔
            if old["stored_filename"]:
                old_path = os.path.join(QUALITY_RAW_DIR, old["stored_filename"])
                if os.path.exists(old_path):
                    try:
                        os.replace(old_path, os.path.join(QUALITY_ARCHIVE_DIR, old["stored_filename"]))
                    except OSError:
                        pass   # 归档失败不影响主流程

        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

    return jsonify({
        "status": "ok",
        "upload_id": upload_id,
        "inspect_date": inspect_date,
        "inspect_date_source": date_source,
        "dept": dept,
        "inspections_count": inspections_count,
        "summary_count": summary_count,
        "superseded_old_id": superseded_old_id,
        "sheet3_present": parsed["sheet3_rules"] is not None,
        "message": (
            f"审核日期 {inspect_date}(来源:{date_source})· "
            f"案例 {inspections_count} 笔 · 汇总 {summary_count} 人"
            + (f" · 覆盖旧上传 {superseded_old_id}" if superseded_old_id else "")
        ),
    }), 200


@app.route("/api/quality/drive_sync_status", methods=["GET"])
def quality_drive_sync_status():
    """Phase 4.4 — 回 Drive 自动同步最近一次状态(由 quality_drive_sync.py 写的状态档)。"""
    status_path = os.path.join(QUALITY_DIR, "drive_sync_status.json")
    if not os.path.exists(status_path):
        return jsonify({"ok": False, "message": "尚未同步"}), 200
    try:
        with open(status_path, encoding="utf-8") as f:
            return jsonify({"ok": True, **json.load(f)}), 200
    except (OSError, ValueError) as e:
        return jsonify({"ok": False, "message": f"状态档读取失败:{e}"}), 200


@app.route("/api/quality/health-data", methods=["GET"])
def quality_health_data():
    """前端预热:active uploads / 笔数 / 各 dept 范围 / 最新审核日 / 最近上传。"""
    conn = get_conn()
    try:
        total_uploads_active = conn.execute(
            "SELECT COUNT(*) AS c FROM quality_uploads WHERE status = 'active'"
        ).fetchone()["c"]
        # 只算 active upload 名下的明细(supersede 教训:查询一律 active filter)
        total_inspections = conn.execute(
            "SELECT COUNT(*) AS c FROM quality_inspections "
            "WHERE upload_id IN (SELECT id FROM quality_uploads WHERE status = 'active')"
        ).fetchone()["c"]
        total_summary = conn.execute(
            "SELECT COUNT(*) AS c FROM quality_summary "
            "WHERE upload_id IN (SELECT id FROM quality_uploads WHERE status = 'active')"
        ).fetchone()["c"]

        depts = {}
        for row in conn.execute(
            """
            SELECT dept, COUNT(*) AS count,
                   MIN(inspect_date) AS earliest, MAX(inspect_date) AS latest
            FROM quality_uploads WHERE status = 'active' GROUP BY dept
            """
        ).fetchall():
            depts[row["dept"]] = {
                "count": row["count"],
                "earliest": row["earliest"],
                "latest": row["latest"],
            }

        agg = conn.execute(
            "SELECT MAX(inspect_date) AS latest_inspect_date, MAX(uploaded_at) AS last_upload_at "
            "FROM quality_uploads WHERE status = 'active'"
        ).fetchone()
    finally:
        conn.close()

    return jsonify({
        "status": "ok",
        "total_uploads_active": total_uploads_active,
        "total_inspections": total_inspections,
        "total_summary": total_summary,
        "depts": depts,
        "latest_inspect_date": agg["latest_inspect_date"],
        "last_upload_at": agg["last_upload_at"],
    })


@app.route("/api/quality/uploads-list", methods=["GET"])
def quality_uploads_list():
    """列出 status='active' 的 uploads metadata(供前端选档管理)。"""
    conn = get_conn()
    try:
        rows = conn.execute(
            """
            SELECT id, original_filename, stored_filename, file_size, md5,
                   inspect_date, inspect_date_source, dept,
                   inspections_count, summary_count,
                   uploaded_by_role, uploaded_by_user, uploaded_at, status
            FROM quality_uploads
            WHERE status = 'active'
            ORDER BY inspect_date DESC, uploaded_at DESC
            """
        ).fetchall()
    finally:
        conn.close()
    return jsonify({"status": "ok", "items": [dict(r) for r in rows]})


# ---------- Phase 4.2:质检查询 ----------
@app.route("/api/quality/data")
def quality_data():
    """主查询:KPI 5 卡 + 明细表 + 错误案例 + 错误类型分布 + 进步榜 + 守门说明。

    range=today|week|last-week|month|quarter|year|ytd|custom
    compare=none|last-week|last-month|last-quarter|last-year|yoy
    dept=all|dx|df(本期实际只有 dx)
    agent=<英文 username>(可选;带则只回该组员资料,给组员自查)
    """
    range_key = request.args.get("range", "month")
    start_date = request.args.get("start_date")
    end_date = request.args.get("end_date")
    compare_key = request.args.get("compare", "none")
    dept_filter = request.args.get("dept", "all")
    agent_filter = request.args.get("agent") or None
    if dept_filter not in ("all", "dx", "df"):
        return jsonify({"status": "error", "message": f"未知 dept: {dept_filter}"}), 400

    try:
        start_iso, end_iso, range_label = quality_queries.resolve_quality_range(
            range_key, start_date, end_date
        )
        compare_range = quality_queries.resolve_compare_range(compare_key, start_iso, end_iso)
    except ValueError as e:
        return jsonify({"status": "error", "message": str(e)}), 400

    conn = get_conn()
    try:
        amap = agent_aliases.build_alias_map(conn)
        current = quality_queries.build_period(
            conn, start_iso, end_iso, dept_filter, amap, full=True, agent_filter=agent_filter
        )
        compare = None
        if compare_range is not None:
            c_start, c_end, c_label = compare_range
            compare = quality_queries.build_period(
                conn, c_start, c_end, dept_filter, amap, full=False, agent_filter=agent_filter
            )
            compare["label"] = c_label
            compare["start"] = c_start
            compare["end"] = c_end
            current["improvements"] = quality_queries.get_improvements(
                conn, start_iso, end_iso, c_start, c_end, dept_filter, amap,
                agent_filter=agent_filter
            )
    finally:
        conn.close()

    return jsonify({
        "status": "ok",
        "range": {"label": range_label, "start": start_iso, "end": end_iso},
        "dept_filter": dept_filter,
        "agent_filter": agent_filter,
        "current": current,
        "compare": compare,
    })


@app.route("/api/quality/rules")
def quality_rules():
    """回最新 active 质检上传的 Sheet 3 评分标准原文摘要(现读原档抽取)。"""
    conn = get_conn()
    try:
        rules = quality_queries.get_latest_rules(conn, QUALITY_RAW_DIR)
    finally:
        conn.close()
    return jsonify({"status": "ok", **rules})


@app.route("/api/quality/monthly_ranking", methods=["GET"])
def quality_monthly_ranking():
    """Phase 13-2 月度排名 endpoint。
    沿用 quality_queries 四件套(resolve_quality_range + _fetch_summary
    + _agg_summary + get_summary_table),公式跟 Excel 官方一致:
        score = pass_rate × 100 = (1 - Σdeduction_sum / Σtotal_messages) × 100
    """
    month = request.args.get("month")  # 例 '2026-06',或省略 = 当月
    dept_filter = request.args.get("dept", "dx")

    # dept 验证
    if dept_filter not in ("all", "dx", "df"):
        return jsonify({"status": "error", "message": "dept 必须 all/dx/df"}), 400

    # min_messages 验证(低访问量门槛,可调;默认 QUALITY_RANKING_MIN_MESSAGES_DEFAULT)
    min_msg_str = request.args.get("min_messages")
    try:
        min_msg = int(min_msg_str) if min_msg_str else QUALITY_RANKING_MIN_MESSAGES_DEFAULT
        if min_msg < 0:
            return jsonify({"status": "error", "message": "min_messages 必须 >= 0"}), 400
    except ValueError:
        return jsonify({"status": "error", "message": "min_messages 必须整数"}), 400

    # month parse → 月份起讫
    if month:
        try:
            y, m = month.split("-")
            y, m = int(y), int(m)
            if not (1 <= m <= 12):
                raise ValueError("month 超出 1-12")
            start_iso = f"{y:04d}-{m:02d}-01"
            # 月底 = 次月 1 号 - 1 天
            nxt = f"{y + 1:04d}-01-01" if m == 12 else f"{y:04d}-{m + 1:02d}-01"
            end_iso = (datetime.datetime.strptime(nxt, "%Y-%m-%d")
                       - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
        except (ValueError, IndexError):
            return jsonify({"status": "error", "message": "month 格式 YYYY-MM"}), 400
    else:
        # 没传 month → 用当月
        start_iso, end_iso, _ = quality_queries.resolve_quality_range("month", None, None)

    conn = get_conn()
    try:
        amap = agent_aliases.build_alias_map(conn)
        canon_of = lambda n: agent_aliases.resolve_name(n, None, amap)  # noqa: E731

        rows = quality_queries._fetch_summary(conn, start_iso, end_iso, dept_filter)
        # _agg_summary 聚合后丢了原始 account;先从 raw rows 建 canon→account 映射补回
        acct_of = {}
        for r in rows:
            c = canon_of(r["agent_name"])
            if r["agent_account"] and c not in acct_of:
                acct_of[c] = r["agent_account"]

        summary_agg = quality_queries._agg_summary(rows, canon_of)
        table = quality_queries.get_summary_table(summary_agg)
        # get_summary_table 是 pass_rate 升序(差的在前);排名要降序(好的第一名)。
        # 显式排序比 reversed() 稳:None(无资料)永远沉底,同分用讯息量多者优先。
        table_desc = sorted(
            table,
            key=lambda r: (
                r.get("pass_rate") is None,
                -(r.get("pass_rate") or 0),
                -(r.get("total_messages") or 0),
                r.get("canonical_name") or "",
            ),
        )

        # 算「月度绩效分」(= pass_rate × 100)+ 按门槛分流主榜 / 低访问量
        ranking = []           # 主榜(讯息 >= min_msg)
        low_volume = []        # 资料不足(讯息 < min_msg)
        rank_counter = 0
        for row in table_desc:
            pr = row.get("pass_rate") or 0
            name = row.get("canonical_name", "")
            total_msg = row.get("total_messages", 0)
            entry = {
                "agent_name": name,
                "agent_account": acct_of.get(name, ""),
                "shift": row.get("shift") or "",
                "total_messages": total_msg,
                "severe_count": row.get("severe_count", 0),
                "medium_count": row.get("medium_count", 0),
                "minor_count": row.get("minor_count", 0),
                "deduction_sum": round(row.get("deduction_sum", 0), 1),
                "pass_rate_pct": round(pr * 100, 2),
                "score": round(pr * 100, 2),  # = pass_rate_pct 同值
            }
            if total_msg >= min_msg:
                rank_counter += 1
                entry["rank"] = rank_counter
                ranking.append(entry)
            else:
                low_volume.append(entry)

        # low_volume 按讯息量降序(table_desc 是合格率序,低量区改看量)
        low_volume.sort(key=lambda x: -x["total_messages"])

        # KPI:本月覆盖天数 / 总严重 / 团队加权合格率 / 第一名
        if dept_filter in ("dx", "df"):
            cov_sql = ("SELECT COUNT(DISTINCT inspect_date) FROM quality_uploads "
                       "WHERE status='active' AND inspect_date >= ? AND inspect_date <= ? "
                       "AND dept = ?")
            cov_args = (start_iso, end_iso, dept_filter)
        else:
            cov_sql = ("SELECT COUNT(DISTINCT inspect_date) FROM quality_uploads "
                       "WHERE status='active' AND inspect_date >= ? AND inspect_date <= ?")
            cov_args = (start_iso, end_iso)
        coverage_days = conn.execute(cov_sql, cov_args).fetchone()[0] or 0

        # KPI 只算主榜(低访问量不影响团队指标)
        total_severe = sum(r["severe_count"] for r in ranking)
        total_msg_sum = sum(r["total_messages"] for r in ranking)
        total_ded = sum(r["deduction_sum"] for r in ranking)
        team_rate = round((1 - total_ded / total_msg_sum) * 100, 2) if total_msg_sum else 0
        top = ranking[0] if ranking else None

        return jsonify({
            "status": "ok",
            "month": start_iso[:7],  # 例 "2026-06"
            "range": {"start": start_iso, "end": end_iso},
            "dept_filter": dept_filter,
            "min_messages": min_msg,
            "kpi": {
                "coverage_days": coverage_days,
                "total_severe": total_severe,
                "team_pass_rate_pct": team_rate,
                "top_agent": {"name": top["agent_name"], "score": top["score"]} if top else None,
            },
            "ranking": ranking,
            "ranking_count": len(ranking),
            "low_volume": low_volume,
            "low_volume_count": len(low_volume),
        })
    finally:
        conn.close()


# ---------- Bug#2a — SYS-02 当月 tab gid 解析 ----------
SYS02_SHEET_ID = "1lKjyN-jDX4IliiNvOehXmyV9dOLIFn1J0syvCLdjWzk"

# 段B 泛化:白名单(防 SSRF/proxy 滥用,只允许已知 Sheet)
SHEET_WHITELIST = {
    "reply": "1lKjyN-jDX4IliiNvOehXmyV9dOLIFn1J0syvCLdjWzk",
    "ops":   "19KeuX9iq7U-ox4IFOIUEVhNMm-kgSJhBNWtsC9h9g40",
    "attendance": "14QWhZnl7DvnUpaKk_JcqjNamahaYOqRVyoqCRV89Ebc",  # Phase 15 出勤排班
}

# Sheet HTML 内每个 tab 的结构(已勘查验证):
#   [<idx>,0,\"<gid>\",[{\"1\":[[0,0,\"<tab名>\"
# tab 命名实测为 '26年X月回复量';解析时容忍多种月份写法。
_SYS02_TAB_RE = re.compile(
    r'\[\d+,0,\\"(\d+)\\",\[\{\\"1\\":\[\[0,0,\\"([^\\]+?)\\"'
)


def _sys02_parse_ym(name):
    """从 tab 名抽 (year, month);容忍 '26年6月回复量'/'2026/06'/'06月' 等变体。"""
    m = re.search(r"(\d{2})年(\d{1,2})月", name)
    if m:
        return (2000 + int(m.group(1)), int(m.group(2)))
    m = re.search(r"(\d{4})[/\-年](\d{1,2})", name)
    if m:
        return (int(m.group(1)), int(m.group(2)))
    m = re.search(r"^0?(\d{1,2})月", name.strip())
    if m:
        return (datetime.datetime.now().year, int(m.group(1)))
    return None


# Phase 15 出勤排班:tab 命名「N月假表YY」(月在前 / YY 在后;容忍「假表」与 YY 间空白,
# 实测有 '6月假表26' 与 '12月假表 25' 两种写法)。无年份后缀者(如 '6月假表')回 None 避免跨年误判。
_ATTENDANCE_TAB_RE = re.compile(r'(\d{1,2})月假表\s*(\d{2})')


def _attendance_parse_ym(name):
    """解析「N月假表YY」格式 tab 名。例:'6月假表26'→(2026,6) / '12月假表 25'→(2025,12)。"""
    m = _ATTENDANCE_TAB_RE.search(name or "")
    if not m:
        return None
    month = int(m.group(1))
    if month < 1 or month > 12:
        return None
    return (2000 + int(m.group(2)), month)


def _resolve_gid_core(sheet_id, prev=False, target_ym=None, parse_fn=None):
    """共用核心:拉 sheet HTML 解析 tab 清单,找出目标月 tab 的 gid。回 (dict, http_status)。
       Args: sheet_id (str)        - Google Sheet ID
             prev (bool)           - True 时找「当月-1 月」的 tab(月环比用),含跨年回退。
             target_ym (tuple|None)- (year, month) 指定任意月;None 维持 now()/prev 行为(零回归)。
             parse_fn (callable)   - tab 名→(年,月) 解析函数;None 用 _sys02_parse_ym(预设)。
       Returns: ({ ok, gid, tab_name, month, year, sheet_id }, 200);
                找不到目标月则 (ok=False + fallback_gid, 200);拉取/解析失败回 (.., 500)。"""
    import requests
    if parse_fn is None:
        parse_fn = _sys02_parse_ym

    if target_ym is None:
        now = datetime.datetime.now()
        target_year, target_month = now.year, now.month
        if prev:
            target_year = target_year if target_month > 1 else target_year - 1
            target_month = (target_month - 1) or 12
    else:
        target_year, target_month = target_ym

    try:
        r = requests.get(
            f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit",
            timeout=15,
        )
    except Exception as e:
        return {"ok": False, "message": f"Sheet 拉取异常: {e}"}, 500

    if r.status_code != 200:
        return {"ok": False, "message": f"Sheet 拉取失败: HTTP {r.status_code}"}, 500

    # 抽所有 (gid, tab名) 配对并解析年月;按 gid 去重
    tabs = []  # [(year, month, gid, name)]
    seen = set()
    for gid, name in _SYS02_TAB_RE.findall(r.text):
        if gid in seen:
            continue
        ym = parse_fn(name)
        if not ym:
            continue
        seen.add(gid)
        tabs.append((ym[0], ym[1], gid, name))

    if not tabs:
        return {"ok": False, "message": "Sheet 内未解析到任何 tab,HTML 结构可能已变"}, 500

    target = next((t for t in tabs if t[0] == target_year and t[1] == target_month), None)
    if target:
        y, mo, gid, name = target
        return {
            "ok": True,
            "gid": gid,
            "tab_name": name,
            "month": mo,
            "year": y,
            "sheet_id": sheet_id,
        }, 200

    latest = sorted(tabs, key=lambda t: (t[0], t[1]), reverse=True)[0]
    return {
        "ok": False,
        "message": f"目标月 ({target_year}/{target_month:02d}) tab 未建立,最新为 {latest[3]}",
        "fallback_gid": latest[2],
        "fallback_tab_name": latest[3],
        "sheet_id": sheet_id,
    }, 200


def _resolve_current_gid(sheet_id, prev=False, target_ym=None, parse_fn=None):
    """对外 Flask 版:回 jsonify Response(沿用旧 caller SYS-02/03 用法,零回归)。"""
    data, code = _resolve_gid_core(sheet_id, prev=prev, target_ym=target_ym, parse_fn=parse_fn)
    return jsonify(data), code


@app.route("/api/current-month-gid", methods=["GET"])
def current_month_gid():
    """泛化版,?sheet=<key> 白名单 reply/ops"""
    key = request.args.get("sheet")
    prev = request.args.get("prev", "").lower() == "true"
    if not key:
        return jsonify({"ok": False, "message": "缺少 sheet 参数"}), 400
    sheet_id = SHEET_WHITELIST.get(key)
    if not sheet_id:
        return jsonify({
            "ok": False,
            "message": f"未知 sheet '{key}',允许: {list(SHEET_WHITELIST.keys())}",
        }), 400
    return _resolve_current_gid(sheet_id, prev=prev)


@app.route("/api/sys02/current-month-gid", methods=["GET"])
def sys02_current_month_gid():
    """旧 endpoint 保留(前端 SYS-02/03 沿用),内部转 reply 表。"""
    return _resolve_current_gid(SHEET_WHITELIST["reply"])


# ════════════════════ Phase 15 出勤排班 ════════════════════

# 状态分类常量(实测 6月假表26):
#   正常类:班 / 休 / 特休 / 例假(不算异常)
#   主异常:迟到 / 早退 / 缺勤 / 事假(计入绩效);裸「迟」= 主管简写迟到、缺分钟数 → 单独标记
#   副异常:病假 / 生理假 / 家庭照顾假(参考)
ATTENDANCE_NORMAL = {'班', '休', '特休', '例假'}
ATTENDANCE_MAIN_ANOMALY_KEYWORDS = ['迟到', '早退', '缺勤', '事假']   # 计入绩效
ATTENDANCE_MINOR_ANOMALY_KEYWORDS = ['病假', '生理假', '家庭照顾假']  # 参考

# 21 名客服按班别配置(权威 key;Sheet 内栏位顺序固定,positional 对应)
ATTENDANCE_AGENTS = {
    '早班': ['贝果', '卡比', '艾娃', '路奇', '果冻', '艾瑞克', '宋江'],
    '中班': ['百川', '沙西米', '卡姆利', '小玥', '鱼丸', '翅膀', '霄霄'],
    '夜班': ['轩轩', '大雄', '咖啡', '小江', '九节狼', '小邱', '当肯'],
}


def _classify_attendance_status(status):
    """分类状态 → (category, type, detail)。
       '班'→('normal','班',None) / '迟到7M'→('main_anomaly','迟到','7M')
       裸'迟'→('main_anomaly','迟(未填分钟)',None) / '病假'→('minor_anomaly','病假',None)
       ''→('blank',None,None) / 其他→('unknown', 原文, None)。"""
    if not status or status.strip() == '':
        return ('blank', None, None)
    s = status.strip()
    if s in ATTENDANCE_NORMAL:
        return ('normal', s, None)
    # 主异常关键字匹配(带分钟/时长 detail)
    for kw in ATTENDANCE_MAIN_ANOMALY_KEYWORDS:
        if s.startswith(kw):
            detail = s[len(kw):].strip() or None
            return ('main_anomaly', kw, detail)
    # 裸「迟」特例(主管简写,未填分钟数;计入主异常但单独标记,方便回 Sheet 补)
    if s == '迟':
        return ('main_anomaly', '迟(未填分钟)', None)
    # 副异常关键字匹配
    for kw in ATTENDANCE_MINOR_ANOMALY_KEYWORDS:
        if s.startswith(kw):
            detail = s[len(kw):].strip() or None
            return ('minor_anomaly', kw, detail)
    # 未识别(可能新格式 → 回报让主管确认)
    return ('unknown', s, None)


def _fetch_attendance_csv_text(sheet_id, gid):
    """抓公开 CSV 原文(positional 解析用)。
       注:不可用 agents_sync.fetch_csv —— 它回 csv.DictReader,会把第 0 行『备注』
       误当表头,且三班并排栏位会被压平;出勤表必须按栏位 index 解析。"""
    import requests
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    resp = requests.get(url, timeout=20)
    resp.raise_for_status()
    resp.encoding = "utf-8"
    text = resp.text
    head = text.lstrip()[:64].lower()
    if head.startswith("<!doctype") or head.startswith("<html"):
        raise RuntimeError("拉到 HTML 而非 CSV(Sheet 可能未公开为『知道连结的人可检视』)")
    return text


def _parse_attendance_csv(csv_text, year, month):
    """parse Sheet CSV(positional)→ 月度汇总 summary + 异常时间线 timeline + kpi + unknowns。
       栏位布局(实测 6月假表26,0-indexed):
         早班 日期 col 0,客服 col 2-8;中班 日期 col 10,客服 col 12-18;夜班 日期 col 20,客服 col 22-28。
       Row 0=备注、Row 1=表头、Row 2+=每天资料(故从 index 2 起读)。"""
    import csv
    import calendar

    rows = list(csv.reader(io.StringIO(csv_text)))
    if len(rows) < 3:
        return None

    shift_layouts = [
        ('早班', 0, 2, ATTENDANCE_AGENTS['早班']),   # (班别, 日期col, 客服起始col, 客服名单)
        ('中班', 10, 12, ATTENDANCE_AGENTS['中班']),
        ('夜班', 20, 22, ATTENDANCE_AGENTS['夜班']),
    ]

    agent_stats = {}   # {(shift, name): {...}}
    timeline = []
    unknowns = []      # 未识别状态(回报让主管确认)

    last_day = calendar.monthrange(year, month)[1]

    for row_idx in range(2, min(2 + last_day, len(rows))):
        row = rows[row_idx]
        if not row or len(row) < 1:
            continue

        for shift_name, date_col, agent_start_col, agent_names in shift_layouts:
            if len(row) <= date_col:
                continue
            date_str = (row[date_col] or '').strip()
            if not date_str or '/' not in date_str:
                continue
            try:
                _m, day_part = date_str.split('/')
                day = int(day_part)
                date_iso = f"{year:04d}-{month:02d}-{day:02d}"
            except (ValueError, IndexError):
                continue

            for i, agent_name in enumerate(agent_names):
                col_idx = agent_start_col + i
                if len(row) <= col_idx:
                    continue
                status = (row[col_idx] or '').strip()
                category, atype, detail = _classify_attendance_status(status)

                key = (shift_name, agent_name)
                if key not in agent_stats:
                    agent_stats[key] = {
                        'shift': shift_name, 'agent_name': agent_name,
                        'shift_days': 0, 'rest_days': 0,
                        'main_count': 0, 'minor_count': 0, 'anomalies': [],
                    }
                stat = agent_stats[key]

                if status == '班':
                    stat['shift_days'] += 1
                elif status in ('休', '特休', '例假'):
                    stat['rest_days'] += 1
                elif category in ('main_anomaly', 'minor_anomaly'):
                    cat_short = 'main' if category == 'main_anomaly' else 'minor'
                    if cat_short == 'main':
                        stat['main_count'] += 1
                    else:
                        stat['minor_count'] += 1
                    event = {'date': date_iso, 'type': atype, 'detail': detail, 'category': cat_short}
                    stat['anomalies'].append(event)
                    timeline.append({
                        'date': date_iso, 'agent_name': agent_name, 'shift': shift_name,
                        'type': atype, 'detail': detail, 'category': cat_short,
                    })
                elif category == 'unknown':
                    unknowns.append({
                        'date': date_iso, 'agent_name': agent_name,
                        'shift': shift_name, 'raw': status,
                    })

    # 组织成各班别 + 排序(主异常多在前)
    summary = {'早班': [], '中班': [], '夜班': []}
    for (shift, agent_name), stat in agent_stats.items():
        summary[shift].append({
            'agent_name': agent_name, 'shift': shift,
            'shift_days': stat['shift_days'], 'rest_days': stat['rest_days'],
            'main_count': stat['main_count'], 'minor_count': stat['minor_count'],
            'anomalies': stat['anomalies'],
        })
    for shift in summary:
        summary[shift].sort(key=lambda x: (-x['main_count'], -x['minor_count']))

    timeline.sort(key=lambda x: x['date'], reverse=True)  # 最新在前

    all_agents = [a for shift_agents in summary.values() for a in shift_agents]
    main_total = sum(a['main_count'] for a in all_agents)
    minor_total = sum(a['minor_count'] for a in all_agents)
    affected = sum(1 for a in all_agents if a['main_count'] > 0 or a['minor_count'] > 0)

    return {
        'summary': summary,
        'timeline': timeline,
        'unknowns': unknowns,
        'kpi': {
            'main_anomaly_total': main_total,
            'minor_anomaly_total': minor_total,
            'affected_agents': affected,
            'total_agents': len(all_agents),
        },
    }


@app.route("/api/attendance/data", methods=["GET"])
def attendance_data():
    """Phase 15 出勤排班月度资料。?month=YYYY-MM(省略=当月)。"""
    month = request.args.get("month")
    if month:
        try:
            y_str, m_str = month.split('-')
            target_ym = (int(y_str), int(m_str))
            if target_ym[1] < 1 or target_ym[1] > 12:
                raise ValueError
        except (ValueError, IndexError):
            return jsonify({"status": "error", "message": "month 格式 YYYY-MM"}), 400
    else:
        now = datetime.datetime.now()
        target_ym = (now.year, now.month)

    sheet_id = SHEET_WHITELIST.get('attendance')
    if not sheet_id:
        return jsonify({"status": "error", "message": "attendance sheet not configured"}), 500

    gid_result, _code = _resolve_gid_core(
        sheet_id, target_ym=target_ym, parse_fn=_attendance_parse_ym,
    )
    if not gid_result.get('ok'):
        return jsonify({
            "status": "error",
            "message": f"无法找到 {target_ym[0]}-{target_ym[1]:02d} 的排班 tab",
            "resolver_result": gid_result,
        }), 404

    try:
        csv_text = _fetch_attendance_csv_text(sheet_id, gid_result['gid'])
    except Exception as e:
        return jsonify({"status": "error", "message": f"fetch CSV failed: {e}"}), 500

    parsed = _parse_attendance_csv(csv_text, target_ym[0], target_ym[1])
    if parsed is None:
        return jsonify({"status": "error", "message": "parse failed"}), 500

    return jsonify({
        "status": "ok",
        "month": f"{target_ym[0]:04d}-{target_ym[1]:02d}",
        "tab_name": gid_result.get('tab_name'),
        "gid": gid_result.get('gid'),
        **parsed,
    })


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000)
