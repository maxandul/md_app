"""
Zentrale Theme/Styling-Datei für das MD-Prozess-Tool.

Diese Datei definiert ein konsistentes, professionelles Design für alle GUI-Komponenten:
- Farbpalette
- Schriftarten und -größen
- Widget-Styles (Notebook, Buttons, Treeviews, Labels, etc.)
- Einheitliche Abstände und Paddings

Verwendung:
    from theme import apply_theme
    apply_theme(root)  # root ist die Tk-Instanz
"""

from tkinter import ttk
import tkinter as tk


# ============================================================================
# FARBPALETTE
# ============================================================================
class Colors:
    """Zentrale Farbdefinitionen für konsistentes Design."""
    
    # Primärfarben (Windows-inspiriert)
    PRIMARY = "#0078d4"           # Windows-Blau
    PRIMARY_DARK = "#005a9e"      # Dunkleres Blau
    PRIMARY_LIGHT = "#4da3e8"     # Helleres Blau
    
    # Hintergrundfarben
    BG_WHITE = "#ffffff"
    BG_LIGHT = "#f3f3f3"          # Sehr helles Grau
    BG_MEDIUM = "#e1e1e1"         # Mittleres Grau
    BG_DARK = "#d0d0d0"           # Dunkles Grau
    
    # Textfarben
    TEXT_PRIMARY = "#000000"
    TEXT_SECONDARY = "#666666"
    TEXT_DISABLED = "#999999"
    
    # Statusfarben
    SUCCESS = "#107c10"           # Grün
    SUCCESS_BG = "#dff6dd"        # Helles Grün
    WARNING = "#ff8c00"           # Orange
    WARNING_BG = "#fff4ce"        # Helles Orange
    ERROR = "#d13438"             # Rot
    ERROR_BG = "#fde7e9"          # Helles Rot
    INFO = "#0078d4"              # Blau
    INFO_BG = "#e6f2fa"           # Helles Blau
    
    # Treeview-Farben
    TREE_HEADER = "#f0f0f0"
    TREE_ROW_EVEN = "#ffffff"
    TREE_ROW_ODD = "#f9f9f9"
    TREE_SELECTED = "#cce8ff"
    TREE_SELECTED_FOCUS = "#0078d4"


# ============================================================================
# SCHRIFTARTEN
# ============================================================================
class Fonts:
    """Zentrale Schriftart-Definitionen."""
    
    FAMILY = "Segoe UI"           # Standard Windows-Schrift
    FAMILY_MONO = "Consolas"      # Monospace für Code/Zahlen
    
    # Größen
    SIZE_SMALL = 8
    SIZE_NORMAL = 9
    SIZE_MEDIUM = 10
    SIZE_LARGE = 11
    SIZE_XLARGE = 12
    SIZE_HEADING = 13
    
    # Vordefinierte Fonts
    NORMAL = (FAMILY, SIZE_NORMAL)
    BOLD = (FAMILY, SIZE_NORMAL, "bold")
    MEDIUM = (FAMILY, SIZE_MEDIUM)
    MEDIUM_BOLD = (FAMILY, SIZE_MEDIUM, "bold")
    LARGE = (FAMILY, SIZE_LARGE)
    LARGE_BOLD = (FAMILY, SIZE_LARGE, "bold")
    HEADING = (FAMILY, SIZE_HEADING, "bold")


# ============================================================================
# PADDING & SPACING
# ============================================================================
class Spacing:
    """Einheitliche Abstände und Paddings."""
    
    # Padding (innen)
    PAD_SMALL = 4
    PAD_NORMAL = 8
    PAD_MEDIUM = 12
    PAD_LARGE = 16
    PAD_XLARGE = 20
    
    # Margins (außen)
    MARGIN_SMALL = 2
    MARGIN_NORMAL = 4
    MARGIN_MEDIUM = 8
    MARGIN_LARGE = 12


# ============================================================================
# THEME ANWENDEN
# ============================================================================
def apply_theme(root: tk.Tk) -> None:
    """
    Wendet das zentrale Theme auf die gesamte Anwendung an.
    
    Args:
        root: Die Hauptanwendung (tk.Tk-Instanz)
    """
    style = ttk.Style(root)
    
    # Theme-Basis (Windows-kompatibel)
    try:
        style.theme_use('vista')  # Windows Vista/7/10 Theme
    except tk.TclError:
        try:
            style.theme_use('clam')  # Fallback
        except tk.TclError:
            pass  # Default Theme verwenden
    
    # ========================================================================
    # NOTEBOOK (Tabs)
    # ========================================================================
    style.configure('TNotebook',
                    background=Colors.BG_LIGHT,
                    borderwidth=0)
    
    style.configure('TNotebook.Tab',
                    padding=[Spacing.PAD_LARGE, Spacing.PAD_MEDIUM],
                    font=Fonts.MEDIUM,
                    background=Colors.BG_MEDIUM,
                    foreground=Colors.TEXT_PRIMARY,
                    borderwidth=1,
                    focuscolor='none')
    
    style.map('TNotebook.Tab',
              background=[('selected', Colors.BG_WHITE),
                         ('active', Colors.BG_DARK)],
              foreground=[('selected', Colors.PRIMARY)],
              font=[('selected', Fonts.MEDIUM_BOLD)],
              expand=[('selected', [1, 1, 1, 0])])  # Tab etwas höher wenn aktiv
    
    # ========================================================================
    # BUTTONS
    # ========================================================================
    style.configure('TButton',
                    padding=[Spacing.PAD_MEDIUM, Spacing.PAD_SMALL],
                    font=Fonts.NORMAL,
                    borderwidth=1,
                    focuscolor='none')
    
    style.map('TButton',
              background=[('active', Colors.BG_DARK),
                         ('pressed', Colors.BG_MEDIUM)])
    
    # Primary Button (hervorgehoben) - Hellblau mit dunklem Text für gute Lesbarkeit
    style.configure('Primary.TButton',
                    padding=[Spacing.PAD_LARGE, Spacing.PAD_NORMAL],  # Größeres Padding
                    font=Fonts.MEDIUM_BOLD,  # Größere, fettere Schrift
                    background=Colors.PRIMARY_LIGHT,  # Hellblauer Hintergrund
                    foreground=Colors.TEXT_PRIMARY,  # Schwarzer Text (gut lesbar)
                    borderwidth=1,
                    relief='raised')  # 3D-Effekt
    
    style.map('Primary.TButton',
              background=[('active', Colors.PRIMARY),  # Dunkler beim Hover
                         ('pressed', Colors.PRIMARY_DARK)],  # Noch dunkler beim Klick
              relief=[('pressed', 'sunken')])
    
    # Danger Button (für kritische Aktionen)
    style.configure('Danger.TButton',
                    padding=[Spacing.PAD_MEDIUM, Spacing.PAD_SMALL],
                    font=Fonts.NORMAL,
                    foreground=Colors.ERROR,
                    borderwidth=1)
    
    style.map('Danger.TButton',
              background=[('active', Colors.ERROR_BG),
                         ('pressed', Colors.ERROR_BG)])
    
    # ========================================================================
    # LABELS
    # ========================================================================
    style.configure('TLabel',
                    font=Fonts.NORMAL,
                    background=Colors.BG_WHITE,
                    foreground=Colors.TEXT_PRIMARY)
    
    # Status-Labels
    style.configure('Success.TLabel',
                    font=Fonts.NORMAL,
                    foreground=Colors.SUCCESS,
                    background=Colors.BG_WHITE)
    
    style.configure('Warning.TLabel',
                    font=Fonts.NORMAL,
                    foreground=Colors.WARNING,
                    background=Colors.BG_WHITE)
    
    style.configure('Error.TLabel',
                    font=Fonts.NORMAL,
                    foreground=Colors.ERROR,
                    background=Colors.BG_WHITE)
    
    style.configure('Info.TLabel',
                    font=Fonts.NORMAL,
                    foreground=Colors.TEXT_SECONDARY,
                    background=Colors.BG_WHITE)
    
    # Heading-Label
    style.configure('Heading.TLabel',
                    font=Fonts.HEADING,
                    foreground=Colors.TEXT_PRIMARY,
                    background=Colors.BG_WHITE)
    
    # Section-Heading (größer, für Abschnittsüberschriften)
    style.configure('SectionHeading.TLabel',
                    font=Fonts.LARGE_BOLD,
                    foreground=Colors.PRIMARY,
                    background=Colors.BG_WHITE)
    
    # Infotext-Label (für erklärende Texte über Tabellen)
    style.configure('InfoText.TLabel',
                    font=Fonts.NORMAL,
                    foreground=Colors.TEXT_SECONDARY,
                    background=Colors.BG_WHITE)
    
    # ========================================================================
    # TREEVIEW (Tabellen)
    # ========================================================================
    style.configure('Treeview',
                    font=Fonts.NORMAL,
                    background=Colors.BG_WHITE,
                    foreground=Colors.TEXT_PRIMARY,
                    fieldbackground=Colors.BG_WHITE,
                    borderwidth=1,
                    relief='solid',
                    rowheight=25)  # Mehr Zeilenhöhe für bessere Lesbarkeit
    
    style.configure('Treeview.Heading',
                    font=Fonts.BOLD,  # Normale Größe, fett
                    background=Colors.TREE_HEADER,  # Heller Hintergrund
                    foreground=Colors.TEXT_PRIMARY,
                    borderwidth=1,
                    relief='flat',  # Flacheres Design, moderner
                    padding=[Spacing.PAD_NORMAL, Spacing.PAD_SMALL])  # Mehr Padding
    
    style.map('Treeview.Heading',
              background=[('active', Colors.BG_DARK)],  # Hover-Effekt
              relief=[('active', 'raised')])  # Leichte Erhebung beim Hover
    
    style.map('Treeview',
              background=[('selected', Colors.TREE_SELECTED_FOCUS)],
              foreground=[('selected', Colors.BG_WHITE)])
    
    # Treeview mit alternierenden Zeilen (muss in Widget selbst konfiguriert werden)
    # Siehe apply_treeview_alternating_colors()
    
    # ========================================================================
    # FRAMES
    # ========================================================================
    style.configure('TFrame',
                    background=Colors.BG_WHITE,
                    borderwidth=0)
    
    style.configure('Card.TFrame',
                    background=Colors.BG_WHITE,
                    borderwidth=1,
                    relief='solid')
    
    # ========================================================================
    # LABELFRAME
    # ========================================================================
    style.configure('TLabelframe',
                    background=Colors.BG_WHITE,
                    borderwidth=1,
                    relief='solid')
    
    style.configure('TLabelframe.Label',
                    font=Fonts.MEDIUM_BOLD,
                    background=Colors.BG_WHITE,
                    foreground=Colors.PRIMARY)  # Blaue Überschrift
    
    # ========================================================================
    # ENTRY & COMBOBOX
    # ========================================================================
    style.configure('TEntry',
                    font=Fonts.NORMAL,
                    borderwidth=1,
                    relief='solid',
                    padding=Spacing.PAD_SMALL)
    
    style.configure('TCombobox',
                    font=Fonts.NORMAL,
                    borderwidth=1,
                    relief='solid',
                    padding=Spacing.PAD_SMALL)
    
    style.map('TCombobox',
              fieldbackground=[('readonly', Colors.BG_WHITE)])
    
    # ========================================================================
    # PROGRESSBAR
    # ========================================================================
    style.configure('TProgressbar',
                    thickness=20,
                    borderwidth=1,
                    background=Colors.PRIMARY,
                    troughcolor=Colors.BG_LIGHT)
    
    # Success Progressbar
    style.configure('Success.Horizontal.TProgressbar',
                    thickness=20,
                    borderwidth=1,
                    background=Colors.SUCCESS,
                    troughcolor=Colors.BG_LIGHT)
    
    # ========================================================================
    # CHECKBUTTON & RADIOBUTTON
    # ========================================================================
    style.configure('TCheckbutton',
                    font=Fonts.NORMAL,
                    background=Colors.BG_WHITE)
    
    style.configure('TRadiobutton',
                    font=Fonts.NORMAL,
                    background=Colors.BG_WHITE)
    
    # ========================================================================
    # SCROLLBAR
    # ========================================================================
    style.configure('Vertical.TScrollbar',
                    background=Colors.BG_MEDIUM,
                    borderwidth=0,
                    arrowsize=12)
    
    style.configure('Horizontal.TScrollbar',
                    background=Colors.BG_MEDIUM,
                    borderwidth=0,
                    arrowsize=12)


# ============================================================================
# HELPER-FUNKTIONEN
# ============================================================================
def apply_treeview_alternating_colors(treeview: ttk.Treeview) -> None:
    """
    Fügt einem Treeview alternierende Zeilenfarben hinzu.
    
    Muss nach dem Befüllen des Treeviews aufgerufen werden!
    
    Args:
        treeview: Das Treeview-Widget
    """
    treeview.tag_configure('evenrow', background=Colors.TREE_ROW_EVEN)
    treeview.tag_configure('oddrow', background=Colors.TREE_ROW_ODD)
    
    # Tags zu bestehenden Items hinzufügen
    for idx, item in enumerate(treeview.get_children()):
        if idx % 2 == 0:
            treeview.item(item, tags=('evenrow',))
        else:
            treeview.item(item, tags=('oddrow',))


def configure_treeview_for_alternating_rows(treeview: ttk.Treeview) -> None:
    """
    Konfiguriert ein Treeview für automatische alternierende Zeilen.
    
    Ruft diese Funktion direkt nach dem Erstellen des Treeviews auf.
    
    Args:
        treeview: Das Treeview-Widget
    """
    treeview.tag_configure('evenrow', background=Colors.TREE_ROW_EVEN)
    treeview.tag_configure('oddrow', background=Colors.TREE_ROW_ODD)


def get_row_tag(index: int) -> str:
    """
    Gibt den Tag-Namen für eine Zeile basierend auf dem Index zurück.
    
    Args:
        index: Der Zeilenindex (0-basiert)
        
    Returns:
        'evenrow' oder 'oddrow'
    """
    return 'evenrow' if index % 2 == 0 else 'oddrow'


# ============================================================================
# THEME-INFORMATIONEN
# ============================================================================
def print_theme_info() -> None:
    """Gibt Informationen über das aktuelle Theme aus (für Debugging)."""
    print("=" * 60)
    print("MD-Prozess-Tool Theme")
    print("=" * 60)
    print(f"Primary Color:    {Colors.PRIMARY}")
    print(f"Success Color:    {Colors.SUCCESS}")
    print(f"Warning Color:    {Colors.WARNING}")
    print(f"Error Color:      {Colors.ERROR}")
    print(f"Font Family:      {Fonts.FAMILY}")
    print(f"Font Size:        {Fonts.SIZE_NORMAL}pt")
    print("=" * 60)

