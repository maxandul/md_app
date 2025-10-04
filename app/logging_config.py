"""
Zentrale Logging-Konfiguration für das MD-Prozess-Tool.

Bietet einheitliches Logging mit verschiedenen Levels und Ausgabeformaten.
"""

import logging
import sys
from pathlib import Path
from datetime import datetime

def setup_logging(log_level: str = "INFO", log_file: str = None) -> logging.Logger:
    """
    Konfiguriert das Logging für die Anwendung.
    
    Args:
        log_level: Logging-Level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        log_file: Optionaler Pfad zur Log-Datei
        
    Returns:
        Konfigurierter Logger
    """
    # Logger erstellen
    logger = logging.getLogger("md_app")
    logger.setLevel(getattr(logging, log_level.upper()))
    
    # Verhindere doppelte Handler
    if logger.handlers:
        return logger
    
    # Formatter für detaillierte Ausgabe
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Console Handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # File Handler (optional)
    if log_file:
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        
        file_handler = logging.FileHandler(log_path, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    
    return logger

def get_logger(name: str = "md_app") -> logging.Logger:
    """Gibt einen Logger für das angegebene Modul zurück."""
    return logging.getLogger(name)
