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

source_path = r"C:\Users\Nutzer\Desktop\Test_Mappe_2.xlsx"

# ---------------------Auslesen der Excel-Tabelle-----------------------

# Erzeugen eines Data-Frames.
df = pd.read_excel(source_path)

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
# df_points = df[["Name der Aufgabe", "Maximal erreichbare Punktzahl der Aufgabe"]]
df_points = df[["Name Aufgabe", "Maximal erreichbare Punktzahl pro Aufgabe"]]

df = df.dropna(subset="Aufgabe")

print(df["Aufgabe"])

# points_per_excercise_max = {}
# for pos, data in df_points.iterrows():
#     if pd.isna(data[0]):  # Falls keine Zahl eingetragen ist, überspringen.
#         pass
#     # Überprüfen, ob die zu erreichende Punktzahl eine Dezimalstelle hat.
#     elif str(data[1])[-1] == "0":
#         # Wenn z.B. 9.0 zu erreichen, dann soll nur 9 in der Tabelle stehen.
#         points_per_excercise_max[data[0]] = int(data[1])  # Eintrag ins Dictionary.
#     else:
#         points_per_excercise_max[data[0]] = data[1]  # Eintrag ins Dictionary.

# # Liste der Aufgaben, mit den Namen, wie sie auch in der Arbeit vergeben wurden:
# exercises = []
# for key in points_per_excercise_max:
#     key = str(key)
#     if len(key) > 1:
#         if key[-2] == ".":
#             key = key[:-1]
#     exercises.append(str(key))

# # Notenspiegel (die Noten stehen in den Spalten 31-37 der Excel-Tabelle)
# overview_of_grades = [int(i) for i in row_two[31:37]]
# average_score = round(row_four[31], 1)

# # Liste erstellen, in der nur die Rohpunkte in Prozent enthalten sind
# df_percent = df[
#     ["Notenschlüssel"]
# ]  # Neues df, in dem nur die Rohpunkte in Prozent relevant sind

# grading_scale_percent = []

# for pos, data in df_percent.iterrows():
#     if pos == 0 or pos > 16:
#         pass
#     else:
#         grading_scale_percent.append(data[0])

# # Liste erstellen, in der nur die zu den Rohpunkten passenden MSS-Noten enthalten sind
# df_mss = df[["Unnamed: 39"]]  # Neues df, in dem nur die MSS-Punkte relevant sind

# grading_scale_mss = []

# for pos, data in df_mss.iterrows():
#     if pos == 0 or pos > 16:
#         pass
#     else:
#         grading_scale_mss.append(data[0])

# # Liste erstellen, in der nur die zu den Rohpunkten passenden Noten in Worten enthalten sind
# df_grades_in_words = df[
#     ["Unnamed: 40"]
# ]  # Neues df, in dem nur die Noten in Worten relevant sind

# grading_scale_words = []

# for pos, data in df_grades_in_words.iterrows():
#     if pos == 0 or pos > 16:
#         pass
#     else:
#         grading_scale_words.append(data[0])
