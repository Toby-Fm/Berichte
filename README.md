# Berichte Automatisierung

Dieses Projekt enthält Skripte zur Automatisierung der Erstellung von Berichten. Es verwendet die Python-Bibliothek `win32com.client`, um Microsoft Word zu öffnen und ein vorhandenes Word-Dokument zu bearbeiten. Das Skript verwendet auch eine Textdatei, um Daten zu lesen, die in das Dokument eingefügt werden sollen.

## Voraussetzungen

- Python 3.10.10
- `win32com.client` Bibliothek
- `pathlib` und `datetime` Bibliotheken (sind in der Standardbibliothek von Python enthalten)

## Installation

1. Stellen Sie sicher, dass Python 3.10.10 auf Ihrem System installiert ist.
2. Installieren Sie die benötigten Python-Bibliotheken mit pip3:

```shell
pip3 install pywin32
```

3. Für MacOS Benutzer, installieren zusätzlich pyobjc:
```shell
$ pip3 install pyobjc
```

## Ausführung
Navigieren im Terminal zu dem Verzeichnis, in dem sich die Skripte befinden (SRC Verzeichnis), und führen das gewünschte Skript aus:
```shell
python BetriebsBerichte.py
```
oder
```shell
python SchulBerichte.py
```
## Info
Beachte, dass die Pfade in den Skripten entsprechend angepasst werden müssen, um auf die spezifischen Dateien und Verzeichnisse zuzugreifen.