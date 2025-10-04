from __future__ import annotations

from pathlib import Path

from tkinter import messagebox
import pandas as pd

from constants import MDConstants
from data_loader import load_employees, load_config
from services.sap_data_service import check_stammdaten


def check_stammdaten(app) -> None:
    """Prüft EXPORT.xlsx, befüllt Prüftabellen und Label im UI."""
    from services.sap_data_service import check_stammdaten as sap_check_stammdaten
    sap_check_stammdaten(app)