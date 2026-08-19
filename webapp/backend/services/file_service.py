import shutil
import os
from pathlib import Path


def copy_file(source: str, destination: str) -> tuple:
    try:
        if not os.path.exists(source):
            return False, f"Archivo origen no existe: {source}"
        dest_dir = os.path.dirname(destination)
        if dest_dir and not os.path.exists(dest_dir):
            os.makedirs(dest_dir, exist_ok=True)
        shutil.copy2(source, destination)
        return True, None
    except Exception as e:
        return False, str(e)


def create_folder(base_path: str, folder_name: str) -> tuple:
    try:
        full_path = os.path.join(base_path, folder_name)
        os.makedirs(full_path, exist_ok=True)
        return True, full_path, None
    except Exception as e:
        return False, None, str(e)


def validate_path(path: str) -> bool:
    return os.path.exists(path)


def clean_filename(name: str) -> str:
    invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
    result = str(name)
    for char in invalid_chars:
        result = result.replace(char, '_')
    return result.strip()
