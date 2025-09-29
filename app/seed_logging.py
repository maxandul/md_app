from pathlib import Path
from datetime import datetime, timedelta
import random

import pandas as pd

try:
    from .simple_tracking import SimpleTrackingSystem
except ImportError:
    from simple_tracking import SimpleTrackingSystem


def seed_logging_data(num_managers: int = 2, employees_per_manager: int = 3,
                      base_rb_year: int | None = None, base_ab_year: int | None = None) -> None:
    """
    Erzeugt Beispiel-Einträge in tracking/md_logging.csv.
    - Legt für num_managers fiktive Vorgesetzte an, je mit employees_per_manager MA.
    - Loggt Versand (Word/PDF/ggf. Feedback erwartet) und markiert teils als empfangen.
    """
    tracking = SimpleTrackingSystem()

    today = datetime.now()
    rb_year = base_rb_year or today.year - 1
    ab_year = base_ab_year or today.year

    for m in range(1, num_managers + 1):
        mgr_pn = f"9{m:03d}"
        mgr_name = f"VG{m} Mustermann"

        # Simpler Feedback nur für ersten Manager
        for e in range(1, employees_per_manager + 1):
            emp_pn = f"4{m:02d}{e:02d}"
            emp_name = f"MA{m}{e} Beispiel"

            # Mische Dokumenttypen: alle haben Rückblick, einige haben Ausblick, selten Probezeit
            doc_types: list[str] = ["rueckblick"]
            if e % 2 == 0:
                doc_types.append("ausblick")
            if e == employees_per_manager and m % 2 == 0:
                doc_types.append("rueckblick_probezeit")

            tracking.log_versand(
                mgr_pn=mgr_pn,
                mgr_name=mgr_name,
                emp_pn=emp_pn,
                emp_name=emp_name,
                doc_types=doc_types,
                rb_year=rb_year,
                ab_year=ab_year,
            )

            # Simuliere ein paar eingegangene Dateien
            num_to_mark = min(random.choice([1, 2]), len(doc_types))
            doc_types_to_mark = random.sample(doc_types, num_to_mark)
            for doc_type in doc_types_to_mark:
                if doc_type == "rueckblick":
                    tracking.mark_received(mgr_pn, emp_pn, "Rückblick Word")
                    tracking.mark_received(mgr_pn, emp_pn, "Rückblick PDF")
                elif doc_type == "ausblick":
                    tracking.mark_received(mgr_pn, emp_pn, "Ausblick Word")
                    tracking.mark_received(mgr_pn, emp_pn, "Ausblick PDF")


if __name__ == "__main__":
    seed_logging_data()


