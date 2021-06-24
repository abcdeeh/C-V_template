from CV_programs import Excel_file_creator,Copy_and_Paste,sheet,paste1
import openpyxl
#Surface area of ​​metal[cm^2]
area="5"
print("Output file`s name")
ands=input()
file_name= ands +".xlsx"
print("sheet's name")
Sheet=input()
Excel_file_creator.create(file_name,Sheet)
Copy_and_Paste.copy(file_name,Sheet,area,ands)

i = 0

while i < 3:
    print("continue? y/n")
    k=input()
    if k=="y":
        print("Next sheet's name")
        Sheet=input()
        sheet.sheet(file_name,Sheet)
        Copy_and_Paste.copy(file_name,Sheet,area,ands)
    if k=="n":
        workbook = openpyxl.load_workbook(filename=file_name)
        workbook.close()
        break
