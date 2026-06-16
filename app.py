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
import os
import re
import sqlite3

from flask import Flask, jsonify, request

import excel_parser

app = Flask(__name__)

# ---------- 设定 ----------
DATA_DIR = "/var/data/refunds"
DB_PATH = os.path.join(DATA_DIR, "db.sqlite")
RAW_DIR = os.path.join(DATA_DIR, "raw")
ARCHIVE_DIR = os.path.join(DATA_DIR, "archive")

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
        "version": "phase-2.2",
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


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000)
