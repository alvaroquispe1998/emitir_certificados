import os
import zipfile
from pathlib import Path


def zip_directory(folder, zip_path):
    folder = Path(folder)
    zip_path = Path(zip_path)
    zip_abs = zip_path.resolve()
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(folder):
            for name in files:
                full_path = Path(root) / name
                if full_path.resolve() == zip_abs:
                    continue
                rel_path = full_path.relative_to(folder)
                zf.write(full_path, rel_path.as_posix())
    return zip_path
