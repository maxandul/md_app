# Theme-Nutzung im MD-Prozess-Tool

## Übersicht

Das zentrale Theme in `theme.py` definiert ein konsistentes, professionelles Design für alle GUI-Komponenten.

## Automatisches Styling

Die meisten Widgets werden **automatisch** gestylt, sobald `apply_theme()` in `main.py` aufgerufen wird:
- ✅ Notebook-Tabs
- ✅ Buttons (Standard)
- ✅ Labels
- ✅ Treeviews
- ✅ Entry/Combobox
- ✅ Frames
- ✅ Progressbars
- ✅ Checkbuttons/Radiobuttons

## Spezielle Button-Styles

### Primary Button (hervorgehoben)
```python
button = ttk.Button(parent, text="Wichtig", style='Primary.TButton')
```

### Danger Button (für kritische Aktionen)
```python
button = ttk.Button(parent, text="Löschen", style='Danger.TButton')
```

## Status-Labels

### Erfolgs-Label (grün)
```python
label = ttk.Label(parent, text="Erfolgreich!", style='Success.TLabel')
```

### Warnung-Label (orange)
```python
label = ttk.Label(parent, text="Warnung!", style='Warning.TLabel')
```

### Fehler-Label (rot)
```python
label = ttk.Label(parent, text="Fehler!", style='Error.TLabel')
```

### Info-Label (grau)
```python
label = ttk.Label(parent, text="Information", style='Info.TLabel')
```

### Überschrift
```python
label = ttk.Label(parent, text="Überschrift", style='Heading.TLabel')
```

### Abschnitts-Überschrift (größer, blau)
```python
label = ttk.Label(parent, text="Abschnitt", style='SectionHeading.TLabel')
```

### Info-Text (für Erklärungen über Tabellen)
```python
label = ttk.Label(parent, text="Erklärungstext", style='InfoText.TLabel', wraplength=800)
```

## Treeview mit alternierenden Zeilen

### Variante 1: Nach dem Befüllen
```python
# Treeview erstellen und befüllen
tree = ttk.Treeview(parent, columns=cols, show="headings")
# ... Items einfügen ...

# Alternierende Farben anwenden
from theme import apply_treeview_alternating_colors
apply_treeview_alternating_colors(tree)
```

### Variante 2: Beim Einfügen
```python
from theme import configure_treeview_for_alternating_rows, get_row_tag

# Treeview konfigurieren
tree = ttk.Treeview(parent, columns=cols, show="headings")
configure_treeview_for_alternating_rows(tree)

# Items mit Tags einfügen
for idx, item in enumerate(data):
    tree.insert("", "end", values=item, tags=(get_row_tag(idx),))
```

## Farben und Schriften direkt nutzen

### Farben
```python
from theme import Colors

label = tk.Label(parent, 
                 text="Custom",
                 fg=Colors.PRIMARY,
                 bg=Colors.BG_LIGHT)
```

### Schriften
```python
from theme import Fonts

label = tk.Label(parent,
                 text="Custom",
                 font=Fonts.LARGE_BOLD)
```

### Abstände
```python
from theme import Spacing

frame.grid(padx=Spacing.PAD_MEDIUM, pady=Spacing.PAD_NORMAL)
```

## Card-Frame (mit Rahmen)

```python
card = ttk.Frame(parent, style='Card.TFrame')
```

## Success Progressbar (grün)

```python
progress = ttk.Progressbar(parent, 
                          mode='determinate',
                          style='Success.Horizontal.TProgressbar')
```

## Theme anpassen

### Neue Farbe hinzufügen
```python
# In theme.py -> Colors Klasse
class Colors:
    # ... bestehende Farben ...
    CUSTOM = "#ff0000"  # Deine Farbe
```

### Neuen Button-Style erstellen
```python
# In theme.py -> apply_theme() Funktion
style.configure('Custom.TButton',
                padding=[12, 6],
                font=Fonts.BOLD,
                foreground=Colors.CUSTOM)
```

### Verwendung
```python
button = ttk.Button(parent, text="Custom", style='Custom.TButton')
```

## Debugging

Theme-Informationen ausgeben:
```python
from theme import print_theme_info
print_theme_info()
```

## Best Practices

1. **Konsistenz**: Verwende die vordefinierten Styles statt eigene zu erstellen
2. **Status-Labels**: Nutze Success/Warning/Error für Statusmeldungen
3. **Primary Buttons**: Nur für die Hauptaktion auf einer Seite
4. **Spacing**: Verwende `Spacing`-Konstanten für einheitliche Abstände
5. **Farben**: Nutze `Colors`-Konstanten statt Hex-Werte direkt

## Beispiel: Kompletter Dialog

```python
from tkinter import ttk
from theme import Colors, Fonts, Spacing

def create_dialog(parent):
    dialog = ttk.Frame(parent, style='Card.TFrame')
    
    # Überschrift
    heading = ttk.Label(dialog, 
                       text="Beispiel-Dialog",
                       style='Heading.TLabel')
    heading.grid(row=0, column=0, columnspan=2,
                 padx=Spacing.PAD_LARGE,
                 pady=Spacing.PAD_LARGE)
    
    # Info
    info = ttk.Label(dialog,
                    text="Informationstext",
                    style='Info.TLabel')
    info.grid(row=1, column=0, columnspan=2,
              padx=Spacing.PAD_LARGE,
              pady=Spacing.PAD_NORMAL)
    
    # Buttons
    btn_cancel = ttk.Button(dialog, text="Abbrechen")
    btn_cancel.grid(row=2, column=0,
                   padx=Spacing.PAD_MEDIUM,
                   pady=Spacing.PAD_LARGE)
    
    btn_ok = ttk.Button(dialog, 
                       text="OK",
                       style='Primary.TButton')
    btn_ok.grid(row=2, column=1,
               padx=Spacing.PAD_MEDIUM,
               pady=Spacing.PAD_LARGE)
    
    return dialog
```

## Support

Bei Fragen oder Problemen: Siehe `theme.py` für die vollständige Implementierung.

