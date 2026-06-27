import json
from pathlib import Path


def load_students(data_file: Path) -> list[dict]:
    if not data_file.exists():
        return []

    with open(data_file, "r", encoding="utf-8") as file:
        return json.load(file)


def save_students(data_file: Path, students: list[dict]) -> None:
    data_file.parent.mkdir(exist_ok=True)
    with open(data_file, "w", encoding="utf-8") as file:
        json.dump(students, file, ensure_ascii=False, indent=2)


def clear_students(data_file: Path) -> None:
    save_students(data_file, [])
