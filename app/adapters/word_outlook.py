from pathlib import Path
import pythoncom
import win32com.client as win32


def fill_word_template(template_path: Path, tag_values: dict, output_path: Path):
    """Füllt Rich-Text-ContentControls anhand .Tag = Schlüssel und speichert als DOCX."""
    pythoncom.CoInitialize()
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(str(template_path))
        for cc in doc.ContentControls:
            tag = (cc.Tag or "").strip()
            if tag and tag in tag_values:
                cc.Range.Text = str(tag_values[tag])
        output_path.parent.mkdir(parents=True, exist_ok=True)
        doc.SaveAs2(str(output_path))
        doc.Close(False)
    finally:
        word.Quit()
        pythoncom.CoUninitialize()


