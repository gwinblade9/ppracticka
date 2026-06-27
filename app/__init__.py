from pathlib import Path
from flask import Flask

BASE_DIR = Path(__file__).resolve().parent.parent


def create_app():
    app = Flask(__name__)
    app.config["SECRET_KEY"] = "change-this-secret-key"
    app.config["BASE_DIR"] = BASE_DIR
    app.config["UPLOAD_FOLDER"] = BASE_DIR / "uploads"
    app.config["GENERATED_FOLDER"] = BASE_DIR / "generated"
    app.config["DATA_FOLDER"] = BASE_DIR / "data"

    # Отдельные JSON-файлы, чтобы таблица удостоверений не смешивалась с таблицей дипломов.
    app.config["CERTIFICATE_DATA_FILE"] = BASE_DIR / "data" / "certificates.json"
    app.config["DIPLOMA_DATA_FILE"] = BASE_DIR / "data" / "diplomas.json"

    # Шаблоны документов.
    app.config["CERTIFICATE_TEMPLATE"] = BASE_DIR / "templates_docs" / "blank_template.docx"
    app.config["DIPLOMA_TEMPLATE"] = BASE_DIR / "templates_docs" / "diploma_template.docx"
    app.config["DIPLOMA_INSERT_TEMPLATE"] = BASE_DIR / "templates_docs" / "diploma_insert_template.docx"
    app.config["DIPLOMA_STATEMENT_TEMPLATE"] = BASE_DIR / "templates_docs" / "diploma_statement_template.docx"

    app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024

    app.config["UPLOAD_FOLDER"].mkdir(exist_ok=True)
    app.config["GENERATED_FOLDER"].mkdir(exist_ok=True)
    app.config["DATA_FOLDER"].mkdir(exist_ok=True)

    from app.routes import main_bp
    app.register_blueprint(main_bp)

    return app
