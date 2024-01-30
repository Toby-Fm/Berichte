#Import der benötigten Bibliotheken
from pathlib import Path
import win32com.client
import datetime  # noqa: F401

# Dateipfad für Word Vorlage festlegen
<<<<<<< HEAD:src/SchulBerichte.py
file_path = r"C:\Users\Documents\Programmierung\Python\Berichte\doc\template_schule.docx"    #ggf. eigenen pfad einfügen  # noqa: E501
=======
file_path = r"C:\Users\template_schule.docx"   
>>>>>>> c1ce68a441217760cc8b39ef1bdbf9cc83696446:Code/SchulBerichte.py

# Öffnet Word und lädt die Vorlage
word = win32com.client.Dispatch("Word.Application")
doc = word.Documents.Open(file_path)               
word.Visible = True  

# Dateipfad für Textdatei festlegen
txt_file_path = r".tx\SchulAufgaben.txt"      

#Lesen der Daten aus der Textdatei
with open(txt_file_path, "r", encoding="utf-8") as f:
    data = f.read()

# Suchen und Ersetzen von Textmarkern im Word-Dokument mit den entsprechenden Daten aus der Textdatei # noqa: E501
word.Selection.Find.ClearFormatting()          

nummer = 0
word.Selection.Find.Execute("<NR>")
word.Selection.Range.Text = nummer

<<<<<<< HEAD:src/SchulBerichte.py
name = "Vorname Nachname"
=======
name = ""
>>>>>>> c1ce68a441217760cc8b39ef1bdbf9cc83696446:Code/SchulBerichte.py
word.Selection.Find.Execute("<NAME>")
word.Selection.Range.Text = name

ausbildungsjahr = 0
word.Selection.Find.Execute("<AJ>")
word.Selection.Range.Text = ausbildungsjahr

date1 = "00.00.0000"
word.Selection.Find.Execute("<DATE1>")
word.Selection.Range.Text = date1

date2 = "00.00.0000"
word.Selection.Find.Execute("<DATE2>")
word.Selection.Range.Text = date2

# Zuordnung der Aufgaben für jeden Tag der Woche aus der Textdatei zu einer Variablen und Erstellung von Zuordnungen zwischen diesen Aufgaben und Nachweisen im Word-Dokument. # noqa: E501
montag_start = data.find("<MONTAG>") + len("<MONTAG>")
montag_end = data.find("<DIENSTAG>") 
montag_aufgaben = data[montag_start:montag_end].strip().split("\n")

dienstag_start = data.find("<DIENSTAG>") + len("<DIENSTAG>")
dienstag_end = data.find("<MITTWOCH>")
dienstag_aufgaben = data[dienstag_start:dienstag_end].strip().split("\n")

mittwoch_start = data.find("<MITTWOCH>") + len("<MITTWOCH>")
mittwoch_end = data.find("<DONNERSTAG>")
mittwoch_aufgaben = data[mittwoch_start:mittwoch_end].strip().split("\n")

donnerstag_start = data.find("<DONNERSTAG>") + len("<DONNERSTAG>")
donnerstag_end = data.find("<FREITAG>")
donnerstag_aufgaben = data[donnerstag_start:donnerstag_end].strip().split("\n")

freitag_start = data.find("<FREITAG>") + len("<FREITAG>")
freitag_end = len(data)
freitag_aufgaben = data[freitag_start:freitag_end].strip().split("\n")


#Für Themen
montag_themen_start = data.find("<MONTAG_THEMEN>") + len("<MONTAG_THEMEN>")
montag_themen_end = data.find("<Dienstag_Themen>")
montag_thema = data[montag_themen_start:montag_themen_end].strip().split("\n")

dienstag_themen_start = data.find("<DIENSTAG_THEMEN>") + len("<DIENSTAG_THEMEN>")
dienstag_themen_ende = len(data)
dienstag_thema = data[dienstag_themen_start:dienstag_themen_ende].strip().split("\n") # noqa: E501

mittwoch_themen_start = data.find("<MITTWOCH_THEMEN>") + len("<MITTWOCH_THEMEN>")
mittwoch_themen_end = len(data)
mittwoch_thema = data[mittwoch_themen_start:mittwoch_themen_end].strip().split("\n") # noqa: E501

donnerstag_themen_start = data.find("<DONNERSTAG_THEMEN>") + len("<DONNERSTAG_THEMEN>")
donnerstag_themen_end = len(data)
donnerstag_thema = data[donnerstag_themen_start:donnerstag_themen_end].strip().split("\n") # noqa: E501

freitag_themen_start = data.find("<FREITAG_THEMEN>") + len("<FREITAG_THEMEN>")
freitag_themen_end = len(data)
freitag_thema = data[freitag_themen_start:freitag_themen_end].strip().split("\n")

#Zuordnen 
montag_fach_mapping = {
    "<FACH1>": montag_aufgaben[0],
    "<FACH2>": montag_aufgaben[1],
    "<FACH3>": montag_aufgaben[2],
    "<FACH4>": montag_aufgaben[3],
    "<FACH5>": montag_aufgaben[4],
    "<FACH6>": montag_aufgaben[5]
}

dienstag_fach_mapping = {
    "<FACH7>": dienstag_aufgaben[0],
    "<FACH8>": dienstag_aufgaben[1],
    "<FACH9>": dienstag_aufgaben[2],
}

mittwoch_fach_mapping = {
    "<FACH13>": mittwoch_aufgaben[0],
    "<FACH14>": mittwoch_aufgaben[1],
    "<FACH15>": mittwoch_aufgaben[2],
    "<FACH16>": mittwoch_aufgaben[3],
    "<FACH17>": mittwoch_aufgaben[4],
    "<FACH18>": mittwoch_aufgaben[5]
}

donnerstag_fach_mapping = {
    "<FACH19>": donnerstag_aufgaben[0],
    "<FACH20>": donnerstag_aufgaben[1],
    "<FACH21>": donnerstag_aufgaben[2]

}

freitag_fach_mapping = {
    "<FACH22>": freitag_aufgaben[0],
    "<FACH23>": freitag_aufgaben[1]
}

#Für Thema
montag_themen_mapping =  {
    "<THEMA1>": montag_thema[0],
    "<THEMA2>": montag_thema[1],
    "<THEMA3>": montag_thema[2]
}

dienstag_themen_mapping = {
    "<THEMA4>": dienstag_thema[0],
    "<THEMA5>": dienstag_thema[1],
}

mittwoch_themen_mapping = {
    "<THEMA8>": mittwoch_thema[0],
    "<THEMA9>": mittwoch_thema[1],
    "<THEMA10>": mittwoch_thema[2]
}

donnerstag_themen_mapping = {
    "<THEMA11>": donnerstag_thema[0],
    "<THEMA12>": donnerstag_thema[1],
    "<THEMA13>": donnerstag_thema[2]
}

freitag_themen_mapping = {
    "<THEMA14>": freitag_thema[0],
    "<THEMA15>": freitag_thema[1],
}

# Suchen und Ersetzen der Abschnitte in der Word-Datei
word.Selection.Find.ClearFormatting()
for fach, aufgabe in montag_fach_mapping.items():
    word.Selection.Find.Execute(fach)
    word.Selection.Range.Text = aufgabe

for fach, aufgabe in dienstag_fach_mapping.items():
    word.Selection.Find.Execute(fach)
    word.Selection.Range.Text = aufgabe

for fach, aufgabe in mittwoch_fach_mapping.items():
    word.Selection.Find.Execute(fach)
    word.Selection.Range.Text = aufgabe

for fach, aufgabe in donnerstag_fach_mapping.items():
    word.Selection.Find.Execute(fach)
    word.Selection.Range.Text = aufgabe

for fach, aufgabe in freitag_fach_mapping.items():
    word.Selection.Find.Execute(fach)
    word.Selection.Range.Text = aufgabe

#Für Thema
for themen, thema in montag_themen_mapping.items():
    word.Selection.Find.Execute(themen)
    word.Selection.Range.Text = thema

for themen, thema in dienstag_themen_mapping.items():
    word.Selection.Find.Execute(themen)
    word.Selection.Range.Text = thema

for themen, thema in mittwoch_themen_mapping.items():
    word.Selection.Find.Execute(themen)
    word.Selection.Range.Text = thema

for themen, thema in donnerstag_themen_mapping.items():
    word.Selection.Find.Execute(themen)
    word.Selection.Range.Text = thema

for themen, thema in freitag_themen_mapping.items():
    word.Selection.Find.Execute(themen)
    word.Selection.Range.Text = thema

<<<<<<< HEAD:src/SchulBerichte.py
path = r"C:\Users\Documents\Berichte\PDF\Bericht vom {} bis {}.pdf".format(date1, date2) # noqa: E501
=======

#Überprüfen ob es die Datei schon gibt. 
path = r"C:\Users\Bericht vom {} bis {}.pdf".format(date1, date2) # noqa: E501
>>>>>>> c1ce68a441217760cc8b39ef1bdbf9cc83696446:Code/SchulBerichte.py

if Path(path).exists():

    print("+--------------------------+")
    print("| Die Datei gibt es schon. |")
    print("+--------------------------+")

    doc.Close(False)
    word.Quit()
    #

else: # Speichern der Word-Datei
<<<<<<< HEAD:src/SchulBerichte.py
    new_file_path = r"C:\Users\Documents\Berichte\Word\Bericht vom {} bis {}.docx".format(date1, date2) # noqa: E501
=======
    new_file_path = r"C:\Users\Berichte\Word\Bericht vom {} bis {}.docx".format(date1, date2) # noqa: E501
>>>>>>> c1ce68a441217760cc8b39ef1bdbf9cc83696446:Code/SchulBerichte.py
    doc.SaveAs(new_file_path)

    if new_file_path:
        print("+------------------------------------------+")
        print("| Word-Dokument wurde erfoglreich erstellt.|")
        print("+------------------------------------------+")

    # Exportieren der Word-Datei als PDF
<<<<<<< HEAD:src/SchulBerichte.py
    pdf_file_path = r"C:\Users\Documents\Berichte\PDF\Bericht vom {} bis {}.pdf".format(date1, date2) # noqa: E501
=======
    pdf_file_path = r"C:\Users\Berichte\PDF\Bericht vom {} bis {}.pdf".format(date1, date2) # noqa: E501
>>>>>>> c1ce68a441217760cc8b39ef1bdbf9cc83696446:Code/SchulBerichte.py
    doc.ExportAsFixedFormat(pdf_file_path, ExportFormat=17, OpenAfterExport=False, OptimizeFor=0) # noqa: E501

    if pdf_file_path:
        print("+----------------------------------------+")
        print("| Word-Dokument wurde zu PDF umgewandelt.|")
        print("+----------------------------------------+")
    
    # Schließen der Word-Datei und Beenden von Word
    doc.Close()
    word.Quit()
