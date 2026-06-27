from pathlib import Path
from flask import Blueprint, current_app, render_template, request, redirect, url_for, flash, send_file
from werkzeug.utils import secure_filename

from app.services.file_reader import read_students_from_file
from app.services.storage import load_students, save_students, clear_students
from app.services.document_generator import generate_documents_docx
from app.services.diploma_generator import generate_diploma_documents_zip

main_bp = Blueprint("main", __name__)

DOC_TYPES = {
    "certificate": {
        "title": "Удостоверения",
        "single_title": "удостоверений",
        "upload_title": "Создание удостоверений",
        "description": "Загрузите таблицу с данными слушателей для формирования удостоверений.",
        "data_config": "CERTIFICATE_DATA_FILE",
        "download_name": "ready_certificates.docx",
    },
    "diploma": {
        "title": "Дипломы",
        "single_title": "дипломов",
        "upload_title": "Создание дипломов",
        "description": "Загрузите файл Итоговая.xlsx: из него берутся ФИО, программа, период, часы, оценки по модулям и дата аттестационной комиссии.",
        "data_config": "DIPLOMA_DATA_FILE",
        "download_name": "Готовые_дипломы.zip",
    },
}


def get_doc_config(doc_type: str) -> dict:
    if doc_type not in DOC_TYPES:
        raise ValueError("Неизвестный тип документа")
    return DOC_TYPES[doc_type]


def get_data_file(doc_type: str) -> Path:
    return current_app.config[get_doc_config(doc_type)["data_config"]]


@main_bp.route("/")
def index():
    return render_template("index.html")


@main_bp.route("/<doc_type>")
def mode_index(doc_type):
    try:
        config = get_doc_config(doc_type)
    except ValueError:
        flash("Такого режима нет")
        return redirect(url_for("main.index"))

    return render_template("upload.html", doc_type=doc_type, config=config)


@main_bp.route("/<doc_type>/upload", methods=["POST"])
def upload_file(doc_type):
    try:
        config = get_doc_config(doc_type)
    except ValueError:
        flash("Такого режима нет")
        return redirect(url_for("main.index"))

    uploaded_file = request.files.get("file")

    if not uploaded_file or uploaded_file.filename == "":
        flash("Выберите файл для загрузки")
        return redirect(url_for("main.mode_index", doc_type=doc_type))

    original_name = uploaded_file.filename
    extension = Path(original_name).suffix.lower()

    if extension not in [".docx", ".xlsx"]:
        flash("Можно загрузить только DOCX или XLSX файл")
        return redirect(url_for("main.mode_index", doc_type=doc_type))

    safe_name = secure_filename(original_name) or f"students{extension}"
    upload_path = current_app.config["UPLOAD_FOLDER"] / f"{doc_type}_{safe_name}"
    uploaded_file.save(upload_path)

    try:
        students = read_students_from_file(upload_path, doc_type=doc_type)
    except Exception as error:
        flash(str(error))
        return redirect(url_for("main.mode_index", doc_type=doc_type))

    save_students(get_data_file(doc_type), students)
    return redirect(url_for("main.preview", doc_type=doc_type))


@main_bp.route("/<doc_type>/preview")
def preview(doc_type):
    try:
        config = get_doc_config(doc_type)
    except ValueError:
        flash("Такого режима нет")
        return redirect(url_for("main.index"))

    students = load_students(get_data_file(doc_type))
    headers = []

    if students:
        headers = [key for key in students[0].keys() if key != "id"]

    return render_template("preview.html", doc_type=doc_type, config=config, students=students, headers=headers)


@main_bp.route("/<doc_type>/edit/<int:student_id>", methods=["GET", "POST"])
def edit_student(doc_type, student_id):
    try:
        config = get_doc_config(doc_type)
    except ValueError:
        flash("Такого режима нет")
        return redirect(url_for("main.index"))

    data_file = get_data_file(doc_type)
    students = load_students(data_file)
    student = next((item for item in students if int(item.get("id")) == student_id), None)

    if not student:
        flash("Строка не найдена")
        return redirect(url_for("main.preview", doc_type=doc_type))

    if request.method == "POST":
        for key in list(student.keys()):
            if key == "id":
                continue
            student[key] = request.form.get(key, "").strip()

        save_students(data_file, students)
        flash("Данные сохранены")
        return redirect(url_for("main.preview", doc_type=doc_type))

    fields = {key: value for key, value in student.items() if key != "id"}
    return render_template("edit.html", doc_type=doc_type, config=config, student=student, fields=fields)


@main_bp.route("/<doc_type>/generate", methods=["POST"])
def generate(doc_type):
    try:
        config = get_doc_config(doc_type)
    except ValueError:
        flash("Такого режима нет")
        return redirect(url_for("main.index"))

    students = load_students(get_data_file(doc_type))

    if not students:
        flash("Сначала загрузите файл с данными")
        return redirect(url_for("main.mode_index", doc_type=doc_type))

    if doc_type == "certificate":
        output_path = generate_documents_docx(
            students=students,
            template_path=current_app.config["CERTIFICATE_TEMPLATE"],
            generated_folder=current_app.config["GENERATED_FOLDER"],
        )
    else:
        output_path = generate_diploma_documents_zip(
            students=students,
            diploma_template=current_app.config["DIPLOMA_TEMPLATE"],
            insert_template=current_app.config["DIPLOMA_INSERT_TEMPLATE"],
            statement_template=current_app.config["DIPLOMA_STATEMENT_TEMPLATE"],
            generated_folder=current_app.config["GENERATED_FOLDER"],
        )

    return send_file(output_path, as_attachment=True, download_name=config["download_name"])


@main_bp.route("/<doc_type>/clear", methods=["POST"])
def clear_current_table(doc_type):
    try:
        get_doc_config(doc_type)
    except ValueError:
        flash("Такого режима нет")
        return redirect(url_for("main.index"))

    answer = request.form.get("answer")
    if answer == "yes":
        clear_students(get_data_file(doc_type))
        flash("Текущая таблица очищена. Можно загрузить новую.")
        return redirect(url_for("main.mode_index", doc_type=doc_type))

    flash("Текущая таблица оставлена без изменений")
    return redirect(url_for("main.preview", doc_type=doc_type))
