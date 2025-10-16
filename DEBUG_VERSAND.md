# Debug-Anleitung für Versand-Kennzeichnung

## Problem 1: Grüne Kennzeichnung funktioniert nicht

### Was prüfen:

1. **Tracking-Datei existiert?**
   - Öffne: `tracking/md_logging_2025.csv` (oder dein aktuelles Jahr)
   - Falls nicht vorhanden: Erst einen Versand durchführen!

2. **Log-Datei prüfen:**
   - Öffne: `tracking/app.log`
   - Suche nach: "Tracking geladen für Jahr"
   - Sollte zeigen: Wie viele VGs mit Einträgen gefunden wurden

3. **VG-PN Format prüfen:**
   - In `md_logging_{jahr}.csv`: Wie sehen die VG-PNs aus?
     - Mit .0: `123456.0`
     - Ohne .0: `123456`
     - Mit Nullen: `000123456`
   - In EXPORT.xlsx: Wie sehen die VG-PNs aus?
   - **Müssen übereinstimmen!**

4. **Anzahl Dokumente prüfen:**
   - In `md_logging_{jahr}.csv`: Zähle Zeilen für einen VG
   - Beispiel: VG "111116" sollte ~12-15 Zeilen haben
   - Filter in Excel: `vg_pn = 111116`
   - Anzahl ≥ 3? → sollte grün sein

### Debug-Output im Log:

Nach App-Start und Klick auf "Aktualisieren" im Massenversand sollte erscheinen:
```
INFO: Tracking geladen für Jahr 2025: 5 VGs mit Einträgen
INFO: VG 111116: 14 Dokumente, versendet=True
INFO: VG 111117: 11 Dokumente, versendet=True
INFO: VG 123456: 0 Dokumente, versendet=False
INFO: VG Müller (111116) wird GRÜN markiert (14 Dokumente)
INFO: VG Weber (111117) wird GRÜN markiert (11 Dokumente)
```

---

## Problem 2: Alternierende Zeilen nach Sortierung

### Was prüfen:

1. **Nach Sortierung:**
   - Klicke auf eine Spaltenüberschrift (z.B. "Nachname")
   - Beobachte: Bleiben die Zeilen korrekt abwechselnd weiß/grau?

2. **Treeview-Sort-Methode:**
   - Die App nutzt `app._bind_treeview_sort()` aus `main.py`
   - Diese delegiert an `bind_treeview_sort()` in `ui_utils.py`
   - Diese sollte `_reapply_alternating_tags()` aufrufen

3. **Manuell testen:**
   - Öffne Python-Konsole
   - Prüfe ob die Funktion existiert:
   ```python
   from app.views.ui_utils import _reapply_alternating_tags
   print(_reapply_alternating_tags)  # Sollte nicht None sein
   ```

### Wenn es nicht funktioniert:

**Mögliche Ursachen:**
1. App wurde nicht neu gestartet nach Code-Änderung
2. Alte .pyc Dateien im Cache
3. Import-Problem

**Lösung:**
```batch
# Cache löschen:
rd /s /q app\__pycache__
rd /s /q app\views\__pycache__
rd /s /q app\services\__pycache__
rd /s /q app\controllers\__pycache__

# App neu starten
python app/main.py
```

---

## Schnelltest:

1. Starte App
2. Öffne `tracking/app.log`
3. Gehe zu Massenversand-Tab
4. Drücke F5 oder wechsle Tab (triggert refresh)
5. Schaue in app.log → sollte Debug-Meldungen zeigen
6. Wenn KEINE Debug-Meldungen: Code wird nicht ausgeführt (Cache-Problem!)

---

## Falls immer noch nichts funktioniert:

Schicke mir:
1. Auszug aus `tracking/app.log` (letzte 50 Zeilen)
2. Erste 5 Zeilen aus `md_logging_2025.csv`
3. Screenshot vom Massenversand-Tab

Dann kann ich gezielt helfen! 🔍

