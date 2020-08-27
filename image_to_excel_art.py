from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfile

def calc_col(n):
    div = int(n / 26)
    mod = n % 26
    if(div > 0):
        return calc_col(div-1)+chr(mod+65)
    if(div == 0):
        return chr(mod+65)

def keep_proportion(large_side, width, height):
    if (width > height):
        return (int((large_side / width) * height), large_side)
    else:
        return (large_side, int((large_side / height) * width))


LARGE_SIDE_RESOLUTION = 200

Tk().withdraw()                 # we don't want a full GUI, so keep the root window from appearing
# show an "Open" dialog box and return the path to the selected file
filename = askopenfilename(
    filetypes=[('PNG', '.png')])

save_name = asksaveasfile(mode='w', filetypes = [('Excel', '.xlsx')], defaultextension='.xlsx')

img = Image.open(filename)
width, height = img.size
n_width, n_height = keep_proportion(LARGE_SIDE_RESOLUTION, width, height)
img = img.resize((n_width, n_height), Image.ANTIALIAS)
rgb_im = img.convert('RGB')
excel_file = Workbook()
sheet = excel_file.active
for i in range(1, n_width):
    sheet.column_dimensions[calc_col(i-1)].width = 3
    for j in range(1, n_height):
        string = '%02x%02x%02x' % rgb_im.getpixel((i, j))
        sheet.cell(row=j, column=i).fill = PatternFill(
            start_color=string, end_color=string, fill_type="solid")
excel_file.save(save_name.name)
