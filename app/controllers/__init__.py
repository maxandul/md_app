"""
Controllers-Package für das MD-Prozess-Tool.

Dieses Package enthält alle Controller-Klassen für die UI-Logik:
- dashboard_controller: Dashboard-Management
- ruecklauf_controller: Rücklauf-Verarbeitung
- stammdaten_controller: Stammdaten-Validierung
- verarbeitung_controller: Dokumentenverarbeitung
- versand_controller: E-Mail-Versand
"""

from .dashboard_controller import (
    refresh_dashboard,
    export_dashboard,
    manual_adjustment
)

from .ruecklauf_controller import (
    scan_real
)

from .stammdaten_controller import (
    check_stammdaten
)

from .verarbeitung_controller import (
    run_full_processing
)

from .versand_controller import (
    send_managers,
    send_selected_employees,
    create_vg_ma_relationship,
    refresh_mgr_table,
    refresh_mgr_table_einzel,
    refresh_vg_list,
    refresh_ma_list,
    update_selection_status,
    preview_managers,
    preview_selected,
    render_mail_preview
)

__all__ = [
    # Dashboard Controller
    "refresh_dashboard",
    "export_dashboard", 
    "manual_adjustment",
    
    # Rücklauf Controller
    "scan_real",
    
    # Stammdaten Controller
    "check_stammdaten",
    
    # Verarbeitung Controller
    "run_full_processing",
    
    # Versand Controller
    "send_managers",
    "send_selected_employees",
    "create_vg_ma_relationship",
    "refresh_mgr_table",
    "refresh_mgr_table_einzel",
    "refresh_vg_list",
    "refresh_ma_list",
    "update_selection_status",
    "preview_managers",
    "preview_selected",
    "render_mail_preview"
]
