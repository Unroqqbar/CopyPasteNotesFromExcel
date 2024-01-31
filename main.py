import openpyxl
import pyperclip
from tkinter import Tk
from tkinter import filedialog, messagebox

root = Tk()
root.withdraw()

workbook_selected = False
while not workbook_selected:
    workbook_location = filedialog.askopenfilename(
        title="Select Workbook",
        filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")]
    )
    if not workbook_location:
        messagebox.showerror("Error", "Please select a File")
        continue
    else:
        workbook_selected = True


workbook_name = workbook_location
workbook = openpyxl.load_workbook(workbook_name, data_only=True)

sheet_name = input("What is the exact name of the Sheet inside Excel?\n")
sheet = workbook[sheet_name]
name_spalte = input("A wéienger Kolonn ass de Virnumm vum Schüler? (A, B, C, ...)\n").upper()
score_spalte = input("A wéienger Kolonn ass d'Note vum Schüler? (A, B, C, ...)\n").upper()
moyenne_spalte = input("A wéienger Kolonn ass d'Moyenne vum Schüler? (A, B, C, ...)\n").upper()
reih = int(input("A Wéienger Reih fänken Donnéen un? (1, 2, 3, ...)\n"))-1

name_list = []
for name in sheet[name_spalte][reih:]:
    name_list.append(name.value)

score_list = []
for score in sheet[score_spalte][reih:]:
    score_list.append(score.value)

moyenne_list = []
for moyenne in sheet[moyenne_spalte][reih:]:
    moyenne_list.append(moyenne.value)

for num in range(len(score_list)):
    text = (f"Gudden Owend {name_list[num]},\n"
            f"Hei schécken ech dir deng Note vun der Prüfung: {score_list[num]} (ob 30)."
            f"\nDeng Moyenne ass {moyenne_list[num]}"
            f"\nConfirméier mir w.e.g., dass dat esou stemmt."
            f"\nFalls de deng Kopie wells kucke kommen, ech sin e Freiden vun 11:50 bis 13:30 am Mediabüro."
            f"\nLéif Gréiss,\nC. Michels")
    print(text)
    pyperclip.copy(text)
    input("\n----- Press Enter after pasting! -----\n")
