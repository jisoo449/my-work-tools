# myapp/routes/__init__.py

from .mail_logs import mail_logs_bp

def register_blueprints(app):
    app.register_blueprint(mail_logs_bp)