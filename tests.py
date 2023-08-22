import pandas as pd
import customtkinter
from docx import Document
from docx.shared import Pt
from docx.shared import Cm
from docx.enum.text import WD_LINE_SPACING
from docx.enum.text import WD_ALIGN_PARAGRAPH
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from PIL import Image, ImageTk

source_path = r"C:\Users\Nutzer\Desktop\Mappe1.xlsx"

# Erzeugen eines Data-Frames.
df = pd.read_excel(source_path)
# df = df.dropna(subset="Aufgabe")

# Erste Datenreihe (also die zweite Zeile des Blattes). Gibt eine Serie zurück.
# Dabei ist der ersten Zeile des Blattes (quasi die Überschrift jeder Spalte)
# immer der Wert der Datenreihe zugeordnet, die hier in der eckigen Klammer steht.
row_one = df.iloc[0]

row_two = df.iloc[1]  # Zweite Zeile der Excel-Tabelle

row_four = df.iloc[3]  # Dritte Zeile der Excel-Tabelle

# Die maximal erreichbare Punktzahl.
# Ist so zu lesen: Gib mir den Wert aus der ersten Datenreihe, der in der Spalte
# mit der Überschrift "Maximal erreichbare Gesamtpunktzahl" steht.
points_max = row_one["Maximal erreichbare Gesamtpunktzahl"]

if str(points_max)[-1] == "0":  # Falls .0, dann soll die Null nicht ausgegeben werden.
    points_max = int(points_max)

# Erstellung eines Dictionary, in dem jeder Aufgabe die zu erreichende Punktzahl zugeordnet wird:
df_points = df[["Name Aufgabe", "Maximal erreichbare Punktzahl pro Aufgabe"]]

points_per_excercise_max = {}

for pos, data in df_points.iterrows():
    if (
        pd.isna(data[0]) or data[0] == 0
    ):  # Falls keine Zahl eingetragen ist, überspringen.
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
    key = str(key)
    if len(key) > 1:
        if key[-2] == ".":
            key = key[:-1]
    exercises.append(str(key))

# Notenspiegel (die Noten stehen in den Spalten 31-37 der Excel-Tabelle)
overview_of_grades = [int(i) for i in row_two[31:37]]
average_score = round(row_four[31], 1)

# Liste erstellen, in der nur die Rohpunkte in Prozent enthalten sind
df_percent = df[
    ["Notenschlüssel"]
]  # Neues df, in dem nur die Rohpunkte in Prozent relevant sind

grading_scale_percent = []

for pos, data in df_percent.iterrows():
    if pos == 0 or pos > 16:
        pass
    else:
        grading_scale_percent.append(data[0])

# Liste erstellen, in der nur die zu den Rohpunkten passenden MSS-Noten enthalten sind
df_mss = df[["Unnamed: 39"]]  # Neues df, in dem nur die MSS-Punkte relevant sind

grading_scale_mss = []

for pos, data in df_mss.iterrows():
    if pos == 0 or pos > 16:
        pass
    else:
        grading_scale_mss.append(data[0])

# Liste erstellen, in der nur die zu den Rohpunkten passenden Noten in Worten enthalten sind
df_grades_in_words = df[
    ["Unnamed: 40"]
]  # Neues df, in dem nur die Noten in Worten relevant sind

grading_scale_words = []

for pos, data in df_grades_in_words.iterrows():
    if pos == 0 or pos > 16:
        pass
    else:
        grading_scale_words.append(data[0])


row_index = 2  # Laufindex, um Daten zu jedem Schüler zu sammeln (Start bei 2, da bei 0 und 1 die die Aufgaben stehen)

for row in range(len(df) - 2):
    # Die nächste Zeile fängt zwei Fälle ab:
    # 1. Am Ende der Schülerliste folgen nur noch Nullen. Hier soll nichts geschehen.
    # 2. Wenn ein Schüler nicht mitgeschrieben hat, steht in jeder Zelle seiner Zeile NaN.
    # Eine solche Zeile muss übersprungen werden, damit es keinen Fehler gibt.
    if df.iloc[row_index, 0] == 0 or pd.isna(df.iloc[row_index, 1]):
        pass
    else:
        # Dictionary, in dem einem Schüler zugeordnet wird, wie viele Punkte er pro Aufgabe erhalten hat:
        name = df.iloc[row_index, 0]  # Name des Schülers der aktuellen Zeile.
        # Nur der Teil der aktuellen Zeile, in dem die Ergebnisse der einzelnen Aufgaben stehen
        student = df.iloc[row_index, 1 : len(exercises) + 1]
        score_per_exercise = {}

        for index, item in enumerate(student):
            grade = str(item)
            if grade[-1] == "0":
                score_per_exercise[exercises[index]] = grade[:-2]
            else:
                score_per_exercise[exercises[index]] = grade

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
# print(student)
print(score_per_exercise)
