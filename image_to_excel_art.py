from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfile

def coluna_excel(n):
    divisao = int(n / 26)
    resto = n % 26
    if(divisao > 0):
        return coluna_excel(divisao-1)+chr(resto+65)
    if(divisao == 0):
        return chr(resto+65)

Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file

save_name = asksaveasfile(mode='a', filetype = [('Excel', '*.xlsx')], defaultextension=[('Excel', '*.xlsx')])

img = Image.open(filename)
img = img.resize((200, 200), Image.ANTIALIAS)
rgb_im = img.convert('RGB')
print(img.size)
width, height = img.size
arquivo_excel = Workbook()
planilha1 = arquivo_excel.active
for i in range(1, width):
    planilha1.column_dimensions[coluna_excel(i-1)].width = 3
    for j in range(1, height):
        string = '%02x%02x%02x' % rgb_im.getpixel((i, j))
        planilha1.cell(row=j, column=i).fill = PatternFill(
            start_color=string, end_color=string, fill_type="solid")
arquivo_excel.save(save_name.name)