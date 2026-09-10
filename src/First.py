# Source: "How to Replace VBA with Python(Step-By-Step Tutorial)" https://www.youtube.com/watch?v=6qo3ly3-_I8&t=457s
import openpyxl
book = openpyxl.load_workbook(r"20210630 NFMW Funds Reg28.xlsx")
sheet=book['PFFB3']
print(sheet['D7'].value)

sheet["b15"] = "Hello there"
book.save(r"20210630 NFMW Funds Reg28.xlsx")

for row in sheet.iter_rows(min_row=1, max_row=5,min_col=3,max_col=6,values_only=True):
    print(row)