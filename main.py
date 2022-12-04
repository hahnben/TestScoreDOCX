import pandas as pd
import customtkinter
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_LINE_SPACING
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from PIL import Image, ImageTk

source_path = r"C:\Users\Nutzer\Desktop\Klausuren_Test_1.xlsx"

# Erzeugen eines Data-Frames.
df = pd.read_excel(source_path)

# Erste Datenreihe (also die zweite Zeile des Blattes). Gibt eine Serie zurück.
# Dabei ist der ersten Zeile des Blattes (quasi die Überschrift jeder Spalte)
# immer der Wert der Datenreihe zugeordnet, die hier in der eckigen Klammer steht.
row_one = df.iloc[0]

# Die maximal erreichbare Punktzahl.
# Ist so zu lesen: Gib mir den Wert aus der ersten Datenreihe, der in der Spalte
# mit der Überschrift "Maximal erreichbare Gesamtpunktzahl" steht.
points_max = row_one["Maximal erreichbare Gesamtpunktzahl"]

# Erstellung eines Dictionary, in dem jeder Aufgabe die zu erreichende Punktzahl zugeordnet wird:
df_points = df[["Name der Aufgabe", "Maximal erreichbare Punktzahl der Aufgabe"]]
points_per_excercise_max = {}
for pos, data in df_points.iterrows():
    if pd.isna(data[0]):  # Falls keine Zahl eingetragen ist, überspringen.
        pass
    # Überprüfen, ob die zu erreichende Punktzahl eine Dezimalstelle hat.
    elif str(data[1])[-1] == "0":
        # Wenn z.B. 9.0 zu erreichen, dann soll nur 9 in der Tabelle stehen.
        points_per_excercise_max[data[0]] = int(data[1])  # Eintrag ins Dictionary.
    else:
        points_per_excercise_max[data[0]] = data[1]  # Eintrag ins Dictionary.


# Liste der Aufgaben, mit den Namen, wie sie auch in der Arbeit vergeben wurden:
exercises = []
for key in points_per_excercise_max:
    exercises.append(key)

document = Document()


def create_student(
    name,
    mss_note,
    note,
    anzahl_aufgaben,
    punkte_pro_aufgabe,
    punkte_aufgaben_erreicht,
    punkte_gesamt,
    erreicht_gesamt,
    prozent,
):
    title = document.add_paragraph()
    title.add_run(name).font.size = Pt(26)
    title.paragraph_format.line_spacing = 1
    title.paragraph_format.space_after = Pt(1)
    line = document.add_paragraph(
        "_________________________________________________________________________________________________________"
    )
    line.paragraph_format.space_before = Pt(1)
    document.add_paragraph("")
    paragraph = document.add_paragraph()
    paragraph.add_run("MSS-Punkte: ").font.size = Pt(16)
    grade = paragraph.add_run(str(mss_note))
    grade.bold = True
    grade.font.size = Pt(16)
    paragraph = document.add_paragraph()
    paragraph.add_run("Note: ").font.size = Pt(16)
    grade = paragraph.add_run(note)
    grade.bold = True
    grade.font.size = Pt(16)
    document.add_paragraph()
    document.add_paragraph().add_run("Punkteverteilung:").font.size = Pt(14)

    # Um die Formatierung in der Tabelle beizubehalten, darf diese nicht zu viele Spalten haben.
    # Als Grenze wird eine Länge von 11 Spalten verwendet. Ist die Tabelle länger, wird sie aufgeteilt.

    if len(exercises) <= 11:
        table = document.add_table(rows=3, cols=anzahl_aufgaben + 2)
        table.style = "Light Shading"
        document.add_paragraph()
        percentage = document.add_paragraph().add_run(
            "Die von dir erreichte Gesamtpunktzahl entspricht "
            + str(prozent)
            + " % der maximal erreichbaren Punkte."
        )
        percentage.font.size = Pt(14)

        for row in table.rows:
            for cell in row.cells:
                if cell == table.rows[0].cells[0]:
                    pass
                else:
                    cell.text = ""
        table.rows[1].cells[0].text = "Punkte maximal"
        table.rows[2].cells[0].text = "Punkte erreicht"

        for i in range(len(table.columns)):
            if i == 0:
                pass
            elif i == (len(table.columns) - 1):
                table.rows[0].cells[i].text = "Gesamt"
            else:
                table.rows[0].cells[i].text = str(exercises[i - 1])

        table_column_index = 1
        for item in punkte_pro_aufgabe:
            table.rows[1].cells[table_column_index].text = str(punkte_pro_aufgabe[item])
            table_column_index += 1

        table_column_index_2 = 1
        for item in punkte_aufgaben_erreicht:
            table.rows[2].cells[table_column_index_2].text = str(
                punkte_aufgaben_erreicht[item]
            )
            table_column_index_2 += 1

        table.rows[1].cells[(len(table.columns) - 1)].text = str(punkte_gesamt)
        table.rows[2].cells[(len(table.columns) - 1)].text = str(erreicht_gesamt)

    else:
        table_1 = document.add_table(rows=3, cols=12)
        table_1.style = "Light Shading"
        document.add_paragraph()
        table_2 = document.add_table(rows=3, cols=anzahl_aufgaben - 9)
        table_2.style = "Light Shading"
        document.add_paragraph()
        percentage = document.add_paragraph().add_run(
            "Die von dir erreichte Gesamtpunktzahl entspricht "
            + str(prozent)
            + " % der maximal erreichbaren Punkte."
        )
        percentage.font.size = Pt(14)

        # Füllen der ersten Tabelle.
        for row in table_1.rows:
            for cell in row.cells:
                if cell == table_1.rows[0].cells[0]:
                    pass
                else:
                    cell.text = ""

        table_1.rows[1].cells[0].text = "Punkte maximal"
        table_1.rows[2].cells[0].text = "Punkte erreicht"

        for i in range(len(table_1.columns)):
            if i == 0:
                pass
            else:
                table_1.rows[0].cells[i].text = str(exercises[i - 1])

        table_column_index = 1
        for index, item in enumerate(punkte_pro_aufgabe):
            if index <= 10:
                table_1.rows[1].cells[table_column_index].text = str(
                    punkte_pro_aufgabe[item]
                )
                table_column_index += 1

        table_column_index_2 = 1
        for index, item in enumerate(punkte_aufgaben_erreicht):
            if index <= 10:
                table_1.rows[2].cells[table_column_index_2].text = str(
                    punkte_aufgaben_erreicht[item]
                )
                table_column_index_2 += 1

        # Füllen der zweiten Tabelle.
        for row in table_2.rows:
            for cell in row.cells:
                if cell == table_2.rows[0].cells[0]:
                    pass
                else:
                    cell.text = ""
        table_2.rows[1].cells[0].text = "Punkte maximal"
        table_2.rows[2].cells[0].text = "Punkte erreicht"

        for i in range(len(table_2.columns)):
            if i == 0:
                pass
            elif i == (len(table_2.columns) - 1):
                table_2.rows[0].cells[i].text = "Gesamt"
            else:
                table_2.rows[0].cells[i].text = str(exercises[i + 10])

        table_column_index = 1
        for index, item in enumerate(punkte_pro_aufgabe):
            if index > 10:
                table_2.rows[1].cells[table_column_index].text = str(
                    punkte_pro_aufgabe[item]
                )
                table_column_index += 1

        table_column_index_2 = 1
        for index, item in enumerate(punkte_aufgaben_erreicht):
            if index > 10:
                table_2.rows[2].cells[table_column_index_2].text = str(
                    punkte_aufgaben_erreicht[item]
                )
                table_column_index_2 += 1

        table_2.rows[1].cells[(len(table_2.columns) - 1)].text = str(punkte_gesamt)
        table_2.rows[2].cells[(len(table_2.columns) - 1)].text = str(erreicht_gesamt)


row_index = 2  # Laufindex, um Daten zu jedem Schüler zu sammeln (Start bei 2, da bei 0 und 1 die die Aufgaben stehen)

for row in range(3):
    if df.iloc[row_index, 0] == 0:
        pass
    else:

        # Dictionary, in dem einem Schüler zugeordnet wird, wie viele Punkte er pro Aufgabe erhalten hat:
        name = df.iloc[row_index, 0]  # Name des Schülers der aktuellen Zeile.
        # Nur der Teil der aktuellen Zeile, in dem die Ergebnisse der einzelnen Aufgaben stehen
        student = df.iloc[row_index, 1 : len(exercises) + 1]
        score_per_exercise = {}

        for index, item in enumerate(student):
            score_per_exercise[exercises[index]] = item

        # Dictionary, in dem einem Schüler die erreichte Gesamtpunktzahl, die Prozent der zu erreichenden Gesamtpunktzahl,
        # die MSS-Punkte, und die Note in Worten zugeordnet werden
        student_2 = df.loc[
            row_index,
            ["Punkte gesamt", "Prozent gesamt", "MSS-Punkte", "Note in Worten"],
        ]
        labels = [
            "Punkte gesamt",
            "Prozent gesamt",
            "MSS-Punkte",
            "Note in Worten",
        ]
        index = 0
        results = {}
        for item in student_2:
            results[labels[index]] = item
            index += 1

        row_index += 1

        create_student(
            name,
            int(results["MSS-Punkte"]),
            results["Note in Worten"],
            len(exercises),
            points_per_excercise_max,
            score_per_exercise,
            points_max,
            results["Punkte gesamt"],
            int(results["Prozent gesamt"]),
        )
        document.add_page_break()  # Sorgt dafür, dass mit jedem Schüler eine neue Seite angefangen wird.

document.save(r"C:\Users\Nutzer\Desktop\Klausurergebnisse.docx")
# document.save(
#     destPath + "//Klausurergebnisse.docx"
# )  # Erstellt am Ende die eigentliche DOCX.
