from pathlib import Path
from app.services.docx_reader import read_docx_table
from app.services.excel_reader import read_excel_table
from app.services.diploma_reader import read_diploma_excel

ALLOWED_EXTENSIONS = {".docx", ".xlsx"}


def read_students_from_file(file_path: Path, doc_type: str = "certificate") -> list[dict]:
    extension = file_path.suffix.lower()

    if extension not in ALLOWED_EXTENSIONS:
        raise ValueError("Можно загрузить только DOCX или XLSX файл")

    if doc_type == "diploma" and extension == ".xlsx":
        return read_diploma_excel(file_path)

    if extension == ".docx":
        return read_docx_table(file_path)

    if extension == ".xlsx":
        return read_excel_table(file_path)

    raise ValueError("Неподдерживаемый формат файла")
