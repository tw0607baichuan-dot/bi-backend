#!/usr/bin/env python3
"""Phase 4.4 — Drive 自动同步质检 Excel(self-contained,镜像 agents_sync.py 范式)

流程:列 Drive 资料夹 → 档名过滤 质检_YYYY-MM-DD.xlsx → 拿 Drive md5Checksum
比对既有 quality_uploads.md5(= byte md5,与手动上传查重同口径)→ 增量下载
→ quality_parser 解析 → 复刻 app.py upload_quality 入库(uploads/inspections/
summary + supersede + 归档)。不改 schema、不动 app.py 路由。

以 www-data 身分跑(cron 用 sudo -u www-data),避免 root 建出占用 DB 的 journal。
"""
import os, sys, io, re, json, hashlib, sqlite3, datetime, traceback

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

sys.path.insert(0, '/opt/bi-backend')
import quality_parser  # 纯函数:resolve_inspect_date + parse_quality_excel

CREDS_PATH          = '/opt/bi-backend/secrets/google_credentials.json'
FOLDER_ID           = '1AUoe0lW_4mw3Z8W9diz10Kb6hU85U-II'
SCOPES              = ['https://www.googleapis.com/auth/drive.readonly']
DB_PATH             = '/var/data/refunds/db.sqlite'   # 质检三表与退费共用此 DB
QUALITY_DIR         = '/var/data/quality'
QUALITY_RAW_DIR     = os.path.join(QUALITY_DIR, 'raw')
QUALITY_ARCHIVE_DIR = os.path.join(QUALITY_DIR, 'archive')
STATUS_FILE         = os.path.join(QUALITY_DIR, 'drive_sync_status.json')
DEPT                = 'dx'           # 与 upload_quality 一致写死(supersede key 之一)
DRIVE_ROLE          = 'drive_auto'   # 不加 source 栏,改记在 uploaded_by_role/user

FILENAME_RE = re.compile(r'^质检_(\d{4}-\d{2}-\d{2})\.xlsx$')


def log(msg):
    print(f"[{datetime.datetime.now().strftime('%F %T')}] {msg}", flush=True)


def _safe_component(s):
    """复刻 app.py:_safe_component,挡路径穿越 / 非法字元。"""
    s = str(s).strip()
    s = re.sub(r"[^\w一-鿿.-]", "_", s)
    return s or "x"


def save_status(payload):
    os.makedirs(QUALITY_DIR, exist_ok=True)
    tmp = STATUS_FILE + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    os.replace(tmp, STATUS_FILE)


def ingest(conn, file_bytes, original_filename, now_iso):
    """复刻 app.py upload_quality 入库段(uploads/inspections/summary + supersede + 归档)。
    回 dict(inspect_date, inspections, summary, superseded_old_id)。可能抛 QualityParseError。"""
    file_md5 = hashlib.md5(file_bytes).hexdigest()
    file_size = len(file_bytes)

    inspect_date, date_source = quality_parser.resolve_inspect_date(
        original_filename, file_bytes, now_iso
    )
    parsed = quality_parser.parse_quality_excel(file_bytes)
    s1 = parsed["sheet1_rows"]
    s2 = parsed["sheet2_rows"]

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    stored_filename = (
        f"{_safe_component(inspect_date)}_{_safe_component(DEPT)}_{timestamp}.xlsx"
    )
    os.makedirs(QUALITY_RAW_DIR, exist_ok=True)
    os.makedirs(QUALITY_ARCHIVE_DIR, exist_ok=True)
    stored_path = os.path.join(QUALITY_RAW_DIR, stored_filename)
    with open(stored_path, "wb") as out:
        out.write(file_bytes)

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
            inspect_date, date_source, DEPT,
            DRIVE_ROLE, DRIVE_ROLE, now_iso,
        ),
    )
    upload_id = cur.lastrowid

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
                upload_id, inspect_date, DEPT, r["case_no"], r["shift"],
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
                upload_id, inspect_date, DEPT, r["shift"], r["agent_name"],
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

    # supersede:同 inspect_date + dept 旧 active → superseded,原档归档
    olds = conn.execute(
        """
        SELECT id, stored_filename FROM quality_uploads
        WHERE inspect_date = ? AND dept = ? AND status = 'active' AND id != ?
        """,
        (inspect_date, DEPT, upload_id),
    ).fetchall()
    superseded_old_id = None
    for old in olds:
        conn.execute(
            "UPDATE quality_uploads SET status = 'superseded', "
            "superseded_by = ?, superseded_at = ? WHERE id = ?",
            (upload_id, now_iso, old["id"]),
        )
        superseded_old_id = old["id"]
        if old["stored_filename"]:
            old_path = os.path.join(QUALITY_RAW_DIR, old["stored_filename"])
            if os.path.exists(old_path):
                try:
                    os.replace(old_path, os.path.join(QUALITY_ARCHIVE_DIR, old["stored_filename"]))
                except OSError:
                    pass  # 归档失败不影响主流程

    return {
        "inspect_date": inspect_date,
        "inspections": inspections_count,
        "summary": summary_count,
        "superseded_old_id": superseded_old_id,
    }


def main():
    now_iso = datetime.datetime.now().isoformat(timespec='seconds')
    try:
        creds = service_account.Credentials.from_service_account_file(CREDS_PATH, scopes=SCOPES)
        svc = build('drive', 'v3', credentials=creds)

        res = svc.files().list(
            q=f"'{FOLDER_ID}' in parents and trashed=false",
            fields='files(id,name,modifiedTime,size,md5Checksum)',
            pageSize=100,
            orderBy='modifiedTime desc',
        ).execute()
        all_files = res.get('files', [])
        log(f"Drive 共 {len(all_files)} 档案")

        valid = []
        for f in all_files:
            m = FILENAME_RE.match(f['name'])
            if m:
                f['inspect_date'] = m.group(1)
                valid.append(f)
        log(f"有效命名 {len(valid)} 档(扣掉 {len(all_files) - len(valid)} 非质检档名)")

        conn = sqlite3.connect(DB_PATH, timeout=15)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA busy_timeout=10000")  # 与线上 gunicorn 并发写,容忍 10s 锁等待

        # 既有 md5 集合(status != 'deleted',与上传端点查重同口径:active+superseded 都算已入过)
        seen_md5 = {
            row["md5"] for row in conn.execute(
                "SELECT md5 FROM quality_uploads WHERE status != 'deleted'"
            )
        }

        synced, skipped, errors = [], [], []
        for f in valid:
            inspect_date = f['inspect_date']
            drive_md5 = f.get('md5Checksum', '')

            if drive_md5 and drive_md5 in seen_md5:
                skipped.append(inspect_date)
                continue

            try:
                log(f"下载 {f['name']} (md5={drive_md5[:8]})")
                request = svc.files().get_media(fileId=f['id'])
                buf = io.BytesIO()
                downloader = MediaIoBaseDownload(buf, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()
                file_bytes = buf.getvalue()

                real_md5 = hashlib.md5(file_bytes).hexdigest()
                if drive_md5 and real_md5 != drive_md5:
                    log(f"⚠️ {f['name']} md5 不符(drive={drive_md5[:8]} real={real_md5[:8]}),仍续行")

                try:
                    r = ingest(conn, file_bytes, f['name'], now_iso)
                    conn.commit()
                except quality_parser.QualityParseError as e:
                    conn.rollback()
                    errors.append({'date': inspect_date, 'error': f'解析失败: {e}'})
                    log(f"✗ {inspect_date} 解析失败: {e}")
                    continue

                seen_md5.add(real_md5)
                synced.append({
                    'date': r['inspect_date'],
                    'inspections': r['inspections'],
                    'summary': r['summary'],
                    'superseded_old_id': r['superseded_old_id'],
                })
                log(
                    f"✓ {r['inspect_date']} 入库 案例{r['inspections']}/汇总{r['summary']}"
                    + (f" 覆盖旧{r['superseded_old_id']}" if r['superseded_old_id'] else "")
                )

            except Exception as e:
                conn.rollback()
                errors.append({'date': inspect_date, 'error': str(e)})
                log(f"✗ {inspect_date} 失败: {e}")
                traceback.print_exc()

        conn.close()

        status = {
            'last_sync_at': now_iso,
            'drive_files_total': len(all_files),
            'valid_files': len(valid),
            'synced': synced,
            'skipped': skipped,
            'errors': errors,
            'success': len(errors) == 0,
        }
        save_status(status)
        log(f"完成:入库 {len(synced)} 跳过 {len(skipped)} 失败 {len(errors)}")

    except Exception as e:
        log(f"FATAL: {e}")
        traceback.print_exc()
        save_status({
            'last_sync_at': now_iso,
            'success': False,
            'fatal_error': str(e),
        })
        sys.exit(1)


if __name__ == '__main__':
    main()
