# /db.py
from flask import current_app, g
from sqlalchemy import create_engine

def init_db(app):
    """
    App 시작 시 1회 호출해서 engine(=pool 포함)을 생성/보관
    """
    app.db_engine = create_engine(
        app.config['db_url'],
        encoding="utf-8",
        pool_pre_ping=True,   # 죽은 커넥션 자동 감지
        pool_recycle=1800,    # (옵션) idle timeout 대비 (초)
        pool_size=5,          # (옵션) 기본 풀 크기
        max_overflow=10       # (옵션) burst 허용
    )

def get_db():
    """
    요청 컨텍스트에서 connection 1개를 재사용
    """
    if 'db_conn' not in g:
        engine = current_app.db_engine
        g.db_conn = engine.connect()
    return g.db_conn

def close_db(e=None):
    """
    요청 종료 시 connection만 close (pool로 반환)
    """
    conn = g.pop('db_conn', None)
    if conn is not None:
        conn.close()