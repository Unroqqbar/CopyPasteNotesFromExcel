import openpyxl
import pyperclip

workbook_name = "database.xlsx"
workbook = openpyxl.load_workbook(workbook_name, data_only=True)
sheet = workbook["SEM 1"]

name_list = []
for name in sheet["B"][2:]:
    name_list.append(name.value)

score_list = []
for score in sheet["E"][2:]:
    score_list.append(score.value)

moyenne_list = []
for moyenne in sheet["L"][2:]:
    moyenne_list.append(moyenne.value)

for num in range(len(score_list)):
    text = (f"Gudden Owend {name_list[num]},\nHei schécken ech dir deng Note vun der Prüfung: {score_list[num]} (ob 30)."
            f"\nDeng Moyenne ass {moyenne_list[num]}"
            f"\nConfirméier mir w.e.g., dass dat esou stemmt."
            f"\nFalls de deng Kopie wells kucke kommen, ech sin e Freiden vun 11:50 bis 13:30 am Mediabüro."
            f"\nLéif Gréiss,\nC. Michels")
    print(text)
    pyperclip.copy(text)
    input("\n----- Press Enter after pasting! -----\n")
