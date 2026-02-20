import os
from typing import Any, Dict, List, Optional

import pymysql
from pymysql.connections import Connection


def _get_conn() -> Connection:
    """
    MariaDB/MySQL connection factory.
    운영에서는 .env(또는 시스템 환경변수)로 주입하는 것을 권장합니다.
    """
    host = os.getenv("DB_HOST", "127.0.0.1")
    port = int(os.getenv("DB_PORT", "3306"))
    user = os.getenv("DB_USER", "innogrid")
    password = os.getenv("DB_PASS", "")
    database = os.getenv("DB_NAME", "MSP_Projects")

    return pymysql.connect(
        host=host,
        port=port,
        user=user,
        password=password,
        database=database,
        charset="utf8mb4",
        cursorclass=pymysql.cursors.DictCursor,
        autocommit=True,
    )


def fetch_all_staff(limit: int = 200) -> List[Dict[str, Any]]:
    """
    Inno_Staff 테이블에서 직원 목록을 가져옵니다.
    """
    sql = """
        SELECT
            inno_staff_id,
            department,
            team,
            inno_staff_name,
            position,
            phone,
            email
        FROM Inno_Staff
        ORDER BY inno_staff_id DESC
        LIMIT %s
    """

    conn = _get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute(sql, (limit,))
            rows = cur.fetchall()
            return rows
    finally:
        conn.close()


def fetch_staff_by_id(staff_id: int) -> Optional[Dict[str, Any]]:
    """
    PK로 1명 조회. 없으면 None 반환.
    """
    sql = """
        SELECT
            inno_staff_id,
            department,
            team,
            inno_staff_name,
            position,
            phone,
            email
        FROM Inno_Staff
        WHERE inno_staff_id = %s
        LIMIT 1
    """

    conn = _get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute(sql, (staff_id,))
            row = cur.fetchone()
            return row
    finally:
        conn.close()


def fetch_staff_by_email(email: str) -> Optional[Dict[str, Any]]:
    """
    이메일로 1명 조회. 로그인/권한 체크 등에 사용.
    """
    sql = """
        SELECT
            inno_staff_id,
            department,
            team,
            inno_staff_name,
            position,
            phone,
            email
        FROM Inno_Staff
        WHERE email = %s
        LIMIT 1
    """

    conn = _get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute(sql, (email,))
            row = cur.fetchone()
            return row
    finally:
        conn.close()
