#Import der benötigten Bibliotheken
import win32com.client
import datetime  # noqa: F401

# Dateipfad für Word Vorlage festlegen
file_path = r"C:\Users\"    #ggf. eigenen pfad einfügen  # noqa: E501

# Öffnet Word und lädt die Vorlage
word = win32com.client.Dispatch("Word.Application")
doc = word.Documents.Open(file_path)               
word.Visible = True  

# Dateipfad für Textdatei festlegen
txt_file_path = r"C:\Users\Documents\Arbeit\wochenplan.txt"       #ggf. eigenen pfad einfügen # noqa: E501

#Lesen der Daten aus der Textdatei
with open(txt_file_path, "r", encoding="utf-8") as f:
    data = f.read()

# Suchen und Ersetzen von Textmarkern im Word-Dokument mit den entsprechenden Daten aus der Textdatei # noqa: E501
word.Selection.Find.ClearFormatting()          

nummer = 73  
word.Selection.Find.Execute("<NR>")
word.Selection.Range.Text = nummer

name = ""
word.Selection.Find.Execute("<NAME>")
word.Selection.Range.Text = name

ausbildungsjahr = 2
word.Selection.Find.Execute("<AJ>")
word.Selection.Range.Text = ausbildungsjahr

date1 = "04.03.2023"
word.Selection.Find.Execute("<DATE1>")
word.Selection.Range.Text = date1

date2 = "10.03.2023"
word.Selection.Find.Execute("<DATE2>")
word.Selection.Range.Text = date2

# Zuordnung der Aufgaben für jeden Tag der Woche aus der Textdatei zu einer Variablen und Erstellung von Zuordnungen zwischen diesen Aufgaben und Nachweisen im Word-Dokument. # noqa: E501
montag_start = data.find("<MONTAG>") + len("<MONTAG>")
montag_end = data.find("<Dienstag>") 
montag_aufgaben = data[montag_start:montag_end].strip().split("\n")

dienstag_start = data.find("<DIENSTAG>") + len("<DIENSTAG>")
dienstag_end = len(data)
dienstag_aufgaben = data[dienstag_start:dienstag_end].strip().split("\n")

mittwoch_start = data.find("<MITTWOCH>") + len("<MITTWOCH>")
mittwoch_end = len(data)
mittwoch_aufgaben = data[mittwoch_start:dienstag_end].strip().split("\n")

donnerstag_start = data.find("<DONNERSTAG>") + len("<DONNERSTAG>")
donnerstag_end = len(data)
donnerstag_aufgaben = data[donnerstag_start:dienstag_end].strip().split("\n")

freitag_start = data.find("<FREITAG>") + len("<FREITAG>")
freitag_end = len(data)
freitag_aufgaben = data[freitag_start:dienstag_end].strip().split("\n")

montag_nachweis_mapping = {
    "<NACHWEIS1>": montag_aufgaben[0],
    "<NACHWEIS2>": montag_aufgaben[1],
    "<NACHWEIS3>": montag_aufgaben[2],
    "<NACHWEIS4>": montag_aufgaben[3]
}

dienstag_nachweis_mapping = {
    "<NACHWEIS5>": dienstag_aufgaben[0],
    "<NACHWEIS6>": dienstag_aufgaben[1],
    "<NACHWEIS7>": dienstag_aufgaben[2],
    "<NACHWEIS8>": dienstag_aufgaben[3]
}

mittwoch_nachweis_mapping = {
    "<NACHWEIS9>": mittwoch_aufgaben[0],
    "<NACHWEIS10>": mittwoch_aufgaben[1],
    "<NACHWEIS11>": mittwoch_aufgaben[2],
    "<NACHWEIS12>": mittwoch_aufgaben[3]
}

donnerstag_nachweis_mapping = {
    "<NACHWEIS13>": donnerstag_aufgaben[0],
    "<NACHWEIS14>": donnerstag_aufgaben[1],
    "<NACHWEIS15>": donnerstag_aufgaben[2],
    "<NACHWEIS16>": donnerstag_aufgaben[3]
}

freitag_nachweis_mapping = {
    "<NACHWEIS17>": freitag_aufgaben[0],
    "<NACHWEIS18>": freitag_aufgaben[1],
    "<NACHWEIS19>": freitag_aufgaben[2],
    "<NACHWEIS20>": freitag_aufgaben[3]
}

# Suchen und Ersetzen der Abschnitte in der Word-Datei
word.Selection.Find.ClearFormatting()
for nachweis, aufgabe in montag_nachweis_mapping.items():
    word.Selection.Find.Execute(nachweis)
    word.Selection.Range.Text = aufgabe

for nachweis, aufgabe in dienstag_nachweis_mapping.items():
    word.Selection.Find.Execute(nachweis)
    word.Selection.Range.Text = aufgabe

for nachweis, aufgabe in mittwoch_nachweis_mapping.items():
    word.Selection.Find.Execute(nachweis)
    word.Selection.Range.Text = aufgabe

for nachweis, aufgabe in donnerstag_nachweis_mapping.items():
    word.Selection.Find.Execute(nachweis)
    word.Selection.Range.Text = aufgabe

for nachweis, aufgabe in freitag_nachweis_mapping.items():
    word.Selection.Find.Execute(nachweis)
    word.Selection.Range.Text = aufgabe

# Speichern der Word-Datei
new_file_path = r"C:\Users\Documents\Arbeit\Fertige_Dokumente\Word\Bericht vom {} bis {}.docx".format(date1, date2)    #ggf. eigenen pfad einfügen # noqa: E501
doc.SaveAs(new_file_path)

# Exportieren der Word-Datei als PDF
pdf_file_path = r"C:\Users\Documents\Arbeit\Fertige_Dokumente\PDF\Bericht vom {} bis {}.pdf".format(date1, date2)      #ggf. eigenen pfad einfügen # noqa: E501
doc.ExportAsFixedFormat(pdf_file_path, ExportFormat=17, OpenAfterExport=False, OptimizeFor=0) # noqa: E501

# Schließen der Word-Datei und Beenden von Word
doc.Close()
word.Quit()
