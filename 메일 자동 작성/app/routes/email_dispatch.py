# myapp/routes/email_dispatch.py

from flask import Blueprint, render_template, request
from sqlalchemy import text
from ..db import get_db

email_dispatches_bp = Blueprint(
    "email_dispatches",
    __name__,
    url_prefix="/email-dispatches"
)

@email_dispatches_bp.route("/mails", methods=["GET"])
def list_email_dispatches():
    """
    메일 발송 로그 목록 조회 페이지
    - 검색(q): recipient/subject 부분일치
    - 상태(status): SUCCESS/FAILED 등
    - 페이지네이션(page)
    """
    q = (request.args.get("q") or "").strip()
    status = (request.args.get("status") or "").strip()
    page = request.args.get("page", default=1, type=int)

    per_page = 20
    if page < 1:
        page = 1
    offset = (page - 1) * per_page

    where = []
    params = {"limit": per_page, "offset": offset}

    if q:
        where.append("(recipient LIKE :q OR subject LIKE :q)")
        params["q"] = f"%{q}%"

    if status:
        where.append("status = :status")
        params["status"] = status

    where_sql = "WHERE " + " AND ".join(where) if where else ""

    conn = get_db()

    # total count
    total = conn.execute(
        text(f"SELECT COUNT(*) AS cnt FROM mail_send_logs {where_sql}"),
        params
    ).scalar() or 0

    # rows
    rows = conn.execute(
        text(f"""
            SELECT
                id,
                recipient,
                subject,
                status,
                created_at,
                provider_message_id,
                error_message
            FROM mail_send_logs
            {where_sql}
            ORDER BY created_at DESC, id DESC
            LIMIT :limit OFFSET :offset
        """),
        params
    ).fetchall()

    total_pages = (total + per_page - 1) // per_page

    # status 필터 드롭다운 후보(원하면 DB distinct로 뽑아도 됨)
    status_options = ["SUCCESS", "FAILED", "PENDING"]

    return render_template(
        "mail_logs/list.html",
        rows=rows,
        q=q,
        status=status,
        status_options=status_options,
        page=page,
        per_page=per_page,
        total=total,
        total_pages=total_pages
    )