"""
Test-Script um die Tag-Konfiguration zu prüfen.
Führe dies aus NACHDEM du die App gestartet hast.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from pathlib import Path
import pandas as pd

# Test 1: Prüfe Tracking-Datei
jahr = 2025  # Ändere das auf dein Jahr
tracking_path = Path(f"tracking/md_logging_{jahr}.csv")

print("=" * 60)
print("TEST: Versand-Kennzeichnung")
print("=" * 60)

if tracking_path.exists():
    print(f"\n✅ Tracking-Datei gefunden: {tracking_path}")
    
    try:
        df = pd.read_csv(tracking_path, sep=";", encoding="utf-8-sig")
        print(f"✅ Datei geladen: {len(df)} Zeilen")
        print(f"✅ Spalten: {list(df.columns)}")
        
        if "vg_pn" in df.columns:
            # Normalisierung
            df["vg_pn"] = df["vg_pn"].astype(str).str.strip()
            df["vg_pn"] = df["vg_pn"].str.replace(r'\.0$', '', regex=True)
            
            # Zähle pro VG
            vg_counts = df.groupby("vg_pn").size().to_dict()
            
            print(f"\n📊 VGs mit Einträgen: {len(vg_counts)}")
            print(f"\n🔍 Top 10 VGs nach Anzahl Dokumente:")
            sorted_vgs = sorted(vg_counts.items(), key=lambda x: x[1], reverse=True)[:10]
            for vg_pn, count in sorted_vgs:
                marker = "🟢 GRÜN" if count >= 3 else "⚪ Normal"
                print(f"   {vg_pn}: {count} Dokumente {marker}")
            
            print(f"\n✅ VGs mit ≥3 Dokumenten (grün): {sum(1 for c in vg_counts.values() if c >= 3)}")
            print(f"✅ VGs mit <3 Dokumenten (normal): {sum(1 for c in vg_counts.values() if c < 3)}")
        else:
            print("❌ Spalte 'vg_pn' nicht gefunden!")
            
    except Exception as e:
        print(f"❌ Fehler beim Laden: {e}")
else:
    print(f"\n❌ Tracking-Datei NICHT gefunden: {tracking_path}")
    print(f"   Pfad: {tracking_path.absolute()}")
    print(f"\n💡 Lösung: Erst einen Versand durchführen, dann wird die Datei erstellt!")

print("\n" + "=" * 60)
print("NÄCHSTE SCHRITTE:")
print("=" * 60)
print("1. Starte die App: python app/main.py")
print("2. Gehe zu 'MD-Versand' → 'Massenversand'")  
print("3. Öffne tracking/app.log während die App läuft")
print("4. Schaue nach Debug-Meldungen:")
print("   - 'Tracking geladen für Jahr ...'")
print("   - 'VG ... wird GRÜN markiert'")
print("5. Falls KEINE Meldungen: Cache löschen!")
print("   rd /s /q app\\__pycache__")
print("   rd /s /q app\\views\\__pycache__")
print("   rd /s /q app\\services\\__pycache__")
print("=" * 60)

