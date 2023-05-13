Dieses Skript verwendet die Python-Bibliothek `win32com.client`, um Microsoft Word zu öffnen und ein vorhandenes Word-Dokument zu bearbeiten.
Das Skript verwendet auch eine Textdatei, um Daten zu lesen, die in das Dokument eingefügt werden sollen.

PythonV 3.10.10 wird verwendet

1. Install template.docx & Wochenplan.txt

2. ggf. "pip install pywin32"

3. Für MacOS "pip install pyobjc" / Nach der Installation, kann man ersetzen durch "from AppKit import NSWorkSpace" 
   Apple ist etwas anders, daher müsste man ggf. den Code etwas anpassen. 