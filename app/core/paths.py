from pathlib import Path
import sys

def project_root() -> Path:
    return Path(__file__).resolve().parents[2]

def data_dir() -> Path:
    p = project_root() / "data"
    p.mkdir(exist_ok=True)
    return p

def db_path() -> Path:
    return data_dir() / "vocab.db"

def words_xlsx_path() -> Path:
    return data_dir() / "words.xlsx"
