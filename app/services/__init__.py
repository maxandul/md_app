"""
Services-Package für das MD-Prozess-Tool.

Dieses Package enthält alle Service-Klassen für die Geschäftslogik:
- document_service: Dokumentenverarbeitung
- email_service: E-Mail-Versand
- outlook_service: Outlook-Integration
- export_service: Export-Funktionen
- dashboard_service: Dashboard-Management
- file_service: Datei-Operationen
- sap_data_service: SAP-Datenverarbeitung
"""

from .document_service import (
    process_docx_folder,
    process_pdfs,
    build_sap_index,
    _strip_accents,
    _name_matches,
    _pn_in_sap,
    _append_processing_log
)

from .email_service import (
    send_managers,
    send_selected_employees,
    preview_managers,
    preview_selected,
    render_mail_preview
)

from .outlook_service import (
    scan_real,
    _get_sender_address,
    process_mail_attachments
)

from .export_service import (
    export_sap_massenupload,
    export_ds_csv,
    generate_export_data
)

from .dashboard_service import (
    refresh_dashboard,
    export_dashboard,
    manual_adjustment,
    get_dashboard_data
)

from .file_service import (
    move_after_processing,
    organize_files,
    cleanup_temp_files
)

from .sap_data_service import (
    check_stammdaten,
    create_vg_ma_relationship,
    refresh_mgr_table,
    refresh_mgr_table_einzel,
    refresh_vg_list,
    refresh_ma_list,
    update_selection_status
)

__all__ = [
    # Document Service
    "process_docx_folder",
    "process_pdfs",
    "build_sap_index",
    "_strip_accents",
    "_name_matches",
    "_pn_in_sap",
    "_append_processing_log",
    
    # Email Service
    "send_managers",
    "send_selected_employees", 
    "preview_managers",
    "preview_selected",
    "render_mail_preview",
    
    # Outlook Service
    "scan_real",
    "_get_sender_address",
    "process_mail_attachments",
    
    # Export Service
    "export_sap_massenupload",
    "export_ds_csv",
    "generate_export_data",
    
    # Dashboard Service
    "refresh_dashboard",
    "export_dashboard", 
    "manual_adjustment",
    "get_dashboard_data",
    
    # File Service
    "move_after_processing",
    "organize_files",
    "cleanup_temp_files",
    
    # SAP Data Service
    "check_stammdaten",
    "create_vg_ma_relationship",
    "refresh_mgr_table",
    "refresh_mgr_table_einzel",
    "refresh_vg_list",
    "refresh_ma_list",
    "update_selection_status"
]
