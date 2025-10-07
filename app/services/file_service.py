"""
File-Service für das MD-Prozess-Tool.

Dieser Service kapselt alle Funktionen zu Datei-Operationen,
einschließlich Verschiebung und Organisation von Dateien.
"""

from __future__ import annotations
from pathlib import Path
import shutil
from typing import Tuple

from app.constants import MDConstants, ProcStatus
from app.data_loader import load_config


def move_after_processing(input_dir: Path, results: list[dict]):
    """
    Verschiebt Dateien je nach Status:
      - OK  -> /ruecklauf/verarbeitet
      - manuell -> /ruecklauf/unverarbeitet/manuell
    Achtung: nur DOCX (hier); PDFs kommen im separaten Schritt.
    """
    import shutil
    
    # Projektwurzel (eine Ebene über 'app') ermitteln und feste Zielordner nutzen
    CFG = load_config()
    ruecklauf_paths = CFG.get("paths", {}).get("ruecklauf", {})
    # Korrektur: Von services/ aus 2 Ebenen hoch zur Root
    base_dir = Path(__file__).parent.parent.parent  # Von services/ -> app/ -> md_app/
    verarbeitet_dir = base_dir / ruecklauf_paths.get("verarbeitet", "../ruecklauf/verarbeitet").lstrip("../")
    manuell_dir = base_dir / ruecklauf_paths.get("manuell", "../ruecklauf/unverarbeitet/manuell").lstrip("../")
    verarbeitet_dir.mkdir(parents=True, exist_ok=True)
    manuell_dir.mkdir(parents=True, exist_ok=True)

    def _unique_path(base_dir: Path, name: str) -> Path:
        target = base_dir / name
        if not target.exists():
            return target
        stem, suffix = Path(name).stem, Path(name).suffix
        i = 1
        while True:
            cand = base_dir / f"{stem} ({i}){suffix}"
            if not cand.exists():
                return cand
            i += 1

    moved_ok = moved_man = 0
    for r in results:
        fname = r.get("file")
        if not fname:
            continue
        src = input_dir / fname
        if not src.exists() or src.suffix.lower() != MDConstants.ALLOWED_EXTENSIONS[0]:
            continue

        if r.get("status") == ProcStatus.OK.value:
            dst = _unique_path(verarbeitet_dir, fname)
            shutil.move(str(src), str(dst))
            moved_ok += 1
            
            # Tracking wurde bereits in document_service.py durchgeführt
        elif r.get("status") in (ProcStatus.MANUELL.value, ProcStatus.PRUEFUNG_NOETIG.value):
            dst = _unique_path(manuell_dir, fname)
            shutil.move(str(src), str(dst))
            moved_man += 1
            
            # Tracking wurde bereits in document_service.py durchgeführt

    return moved_ok, moved_man


def organize_files(source_dir: Path, target_dir: Path, file_pattern: str = "*") -> int:
    """
    Organisiert Dateien nach bestimmten Kriterien.
    
    Returns:
        Anzahl der organisierten Dateien
    """
    # Optional: Kann später implementiert werden
    pass


def cleanup_temp_files(temp_dir: Path) -> int:
    """
    Bereinigt temporäre Dateien.
    
    Returns:
        Anzahl der gelöschten Dateien
    """
    # Optional: Kann später implementiert werden
    pass
