# app/__init__.py

from flask import Flask
from .db import init_db, close_db
from .routes import bp as main_bp

def create_app():
    app = Flask(__name__)
    app.config.from_pyfile("config.py")

    # DB 엔진 1회 생성
    init_db(app)

    # 요청 종료 시 connection 반환
    app.teardown_appcontext(close_db)

    # Blueprint 등록
    from .routes.users import users_bp
    app.register_blueprint(users_bp)

    return app