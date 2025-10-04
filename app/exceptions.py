"""
Zentrale Exception-Klassen für das MD-Prozess-Tool.

Bietet spezifische Exception-Typen für verschiedene Fehlerfälle.
"""

class MDAppException(Exception):
    """Basis-Exception für alle MD-App-spezifischen Fehler."""
    pass

class ConfigurationError(MDAppException):
    """Fehler bei der Konfiguration."""
    pass

class DataValidationError(MDAppException):
    """Fehler bei der Datenvalidierung."""
    pass

class DocumentProcessingError(MDAppException):
    """Fehler bei der Dokumentenverarbeitung."""
    pass

class EmailServiceError(MDAppException):
    """Fehler beim E-Mail-Versand."""
    pass

class SAPDataError(MDAppException):
    """Fehler bei SAP-Datenverarbeitung."""
    pass

class FileServiceError(MDAppException):
    """Fehler bei Dateioperationen."""
    pass
