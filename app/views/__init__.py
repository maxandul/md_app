"""
Views-Package für das MD-Prozess-Tool.

Dieses Package enthält alle View-Klassen für die UI-Darstellung:
- dashboard_view: Dashboard-UI
- ruecklauf_view: Rücklauf-UI
- stammdaten_view: Stammdaten-UI
- verarbeitung_view: Verarbeitungs-UI
- versand_view: Versand-UI
- ui_utils: UI-Hilfsfunktionen
"""

from .dashboard_view import build_dashboard
from .ruecklauf_view import build_ruecklauf
from .stammdaten_view import build_stammdaten
from .verarbeitung_view import build_verarbeitung
from .versand_view import build_versand
from .ui_utils import (
    make_tree,
    bind_treeview_sort,
    autosize_tree_columns
)

__all__ = [
    # View Builders
    "build_dashboard",
    "build_ruecklauf", 
    "build_stammdaten",
    "build_verarbeitung",
    "build_versand",
    
    # UI Utils
    "make_tree",
    "bind_treeview_sort",
    "autosize_tree_columns"
]
