import win32com.client as win32
import os
from tkinter import Tk, filedialog
from colorama import init, Fore, Style
from datetime import datetime

init(autoreset=True)


def log_action(message):
    with open("log.txt", "a") as log_file:
        log_file.write(f"[{datetime.now()}] {message}\n")


print(Fore.CYAN + r"""
    _    ____  _     _______   __
   / \  |  _ \| |   | ____\ \ / /
  / _ \ | |_) | |   |  _|  \ V / 
 / ___ \|  _ <| |___| |___  | |  
/_/   \_\_| \_\_____|_____| |_|   
""" + Style.RESET_ALL)

print(Fore.GREEN + "Select Package:" + Style.RESET_ALL)
print(Fore.GREEN + "A - 2pcs 2x2, 6pcs 1x1 " + Style.RESET_ALL)
print(Fore.GREEN + "B - 4pcs 2x2, 8pcs 1x1 " + Style.RESET_ALL)
print(Fore.GREEN + "C - 6pcs 2x2, 12pcs 1x1 " + Style.RESET_ALL)
print(Fore.GREEN + "O - Custom Package (custom 1x1, 2x2, passport size)" + Style.RESET_ALL)

package = input(Fore.YELLOW + "Enter Package (A/B/C/O): " + Style.RESET_ALL).strip().lower()

if package not in ['a', 'b', 'c', 'o']:
    print(Fore.RED + "Invalid package. Exiting." + Style.RESET_ALL)
    exit()

package = package.upper()

Tk().withdraw()
image_path = filedialog.askopenfilename(
    title="Select Your ID Picture",
    filetypes=[("Image Files", "*.jpg *.png *.jpeg")]
)

if not image_path:
    print("No file selected. Exiting.")
    exit()

image_path = os.path.abspath(image_path)

word = win32.Dispatch('Word.Application')
word.Visible = True

doc = word.Documents.Add()

# A4
doc.PageSetup.PageWidth = 595.3  # 21cm
doc.PageSetup.PageHeight = 841.9  # 29.7cm

# custom margins
doc.PageSetup.TopMargin = 7.2  # .1 x 72
doc.PageSetup.BottomMargin = 0
doc.PageSetup.LeftMargin = 7.2  # ^^^^
doc.PageSetup.RightMargin = 0

size_2x2 = 144  # 2 inches
size_1x1 = 72  # 1 inch

left_start = 0
top_start = 0
gray_color = 150 + (150 * 256) + (150 * 256 * 256)

if package == 'A':
    for i in range(2):
        left_pos = left_start + i * size_2x2
        shape = doc.Shapes.AddPicture(FileName=image_path, LinkToFile=False, SaveWithDocument=True,
                                      Left=left_pos, Top=top_start,
                                      Width=size_2x2, Height=size_2x2)
        shape.WrapFormat.Type = 3
        shape.Line.Weight = 1
        shape.Line.ForeColor.RGB = gray_color

    left_start_1x1 = left_start + 2 * size_2x2
    for i in range(3):
        left_pos = left_start_1x1 + i * size_1x1
        shape = doc.Shapes.AddPicture(FileName=image_path, LinkToFile=False, SaveWithDocument=True,
                                      Left=left_pos, Top=top_start,
                                      Width=size_1x1, Height=size_1x1)
        shape.WrapFormat.Type = 3
        shape.Line.Weight = 1
        shape.Line.ForeColor.RGB = gray_color

    for i in range(3):
        left_pos = left_start_1x1 + i * size_1x1
        top_pos = top_start + size_1x1
        shape = doc.Shapes.AddPicture(FileName=image_path, LinkToFile=False, SaveWithDocument=True,
                                      Left=left_pos, Top=top_pos,
                                      Width=size_1x1, Height=size_1x1)
        shape.WrapFormat.Type = 3
        shape.Line.Weight = 1
        shape.Line.ForeColor.RGB = gray_color

elif package == 'B':
    for i in range(4):
        left_pos = left_start + i * size_2x2
        shape = doc.Shapes.AddPicture(FileName=image_path, LinkToFile=False, SaveWithDocument=True,
                                      Left=left_pos, Top=top_start,
                                      Width=size_2x2, Height=size_2x2)
        shape.WrapFormat.Type = 3
        shape.Line.Weight = 1
        shape.Line.ForeColor.RGB = gray_color

    top_pos_1x1 = top_start + size_2x2
    for i in range(8):
        left_pos = left_start + i * size_1x1
        shape = doc.Shapes.AddPicture(FileName=image_path, LinkToFile=False, SaveWithDocument=True,
                                      Left=left_pos, Top=top_pos_1x1,
                                      Width=size_1x1, Height=size_1x1)
        shape.WrapFormat.Type = 3
        shape.Line.Weight = 1
        shape.Line.ForeColor.RGB = gray_color

elif package == 'C':
    for i in range(4):
        left_pos = left_start + i * size_2x2
        shape = doc.Shapes.AddPicture(FileName=image_path, LinkToFile=False, SaveWithDocument=True,
                                      Left=left_pos, Top=top_start,
                                      Width=size_2x2, Height=size_2x2)
        shape.WrapFormat.Type = 3
        shape.Line.Weight = 1
        shape.Line.ForeColor.RGB = gray_color

    top_pos_row2 = top_start + size_2x2
    for i in range(2):
        left_pos = left_start + i * size_2x2
        shape = doc.Shapes.AddPicture(FileName=image_path, LinkToFile=False, SaveWithDocument=True,
                                      Left=left_pos, Top=top_pos_row2,
                                      Width=size_2x2, Height=size_2x2)
        shape.WrapFormat.Type = 3
        shape.Line.Weight = 1
        shape.Line.ForeColor.RGB = gray_color

    left_start_1x1 = left_start + 2 * size_2x2
    for i in range(4):
        left_pos = left_start_1x1 + i * size_1x1
        shape = doc.Shapes.AddPicture(FileName=image_path, LinkToFile=False, SaveWithDocument=True,
                                      Left=left_pos, Top=top_pos_row2,
                                      Width=size_1x1, Height=size_1x1)
        shape.WrapFormat.Type = 3
        shape.Line.Weight = 1
        shape.Line.ForeColor.RGB = gray_color

    top_pos_row3 = top_pos_row2 + size_1x1
    left_start_row3 = left_start + 2 * size_2x2
    for i in range(4):
        left_pos = left_start_row3 + i * size_1x1
        shape = doc.Shapes.AddPicture(FileName=image_path, LinkToFile=False, SaveWithDocument=True,
                                      Left=left_pos, Top=top_pos_row3,
                                      Width=size_1x1, Height=size_1x1)
        shape.WrapFormat.Type = 3
        shape.Line.Weight = 1
        shape.Line.ForeColor.RGB = gray_color

    top_pos_row4 = top_pos_row3 + size_1x1
    for i in range(4):
        left_pos = left_start + i * size_1x1
        shape = doc.Shapes.AddPicture(FileName=image_path, LinkToFile=False, SaveWithDocument=True,
                                      Left=left_pos, Top=top_pos_row4,
                                      Width=size_1x1, Height=size_1x1)
        shape.WrapFormat.Type = 3
        shape.Line.Weight = 1
        shape.Line.ForeColor.RGB = gray_color

elif package == 'O':
    print("\nYou selected CUSTOM PACKAGE (PACKAGE O)\n")

    print(Fore.CYAN + "Choose Picture Size:" + Style.RESET_ALL)
    print(Fore.CYAN + "1 - 1x1 in" + Style.RESET_ALL)
    print(Fore.CYAN + "2 - 2x2 in" + Style.RESET_ALL)
    print(Fore.CYAN + "P - Passport size (1.4in x 1.8in)" + Style.RESET_ALL)

    pic_size_input = input(Fore.YELLOW + "Enter picture size (1/2/P): " + Style.RESET_ALL).strip().lower()

    if pic_size_input == '2':
        pic_width = 144  # 2in
        pic_height = 144
        pic_label = "2x2"
    elif pic_size_input == '1':
        pic_width = 72  # 1in
        pic_height = 72
        pic_label = "1x1"
    elif pic_size_input == 'p':
        pic_width = 100.8  # 1.4in 72 * 1.4
        pic_height = 129.6  # 1.8in 72 * 1.8
        pic_label = "Passport"
    else:
        print(Fore.RED + "Invalid size selected. Exiting." + Style.RESET_ALL)
        exit()

    try:
        quantity = int(input(f"How many {pic_label} pictures do you want? "))
    except ValueError:
        print(Fore.RED + "Invalid number. Exiting." + Style.RESET_ALL)
        exit()

    log_action(f"User selected Custom Package: {quantity} pcs {pic_label}")

    max_width = 595.3 - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin
    pics_per_row = int(max_width // pic_width)

    left_pos = left_start
    top_pos = top_start

    for i in range(quantity):
        shape = doc.Shapes.AddPicture(FileName=image_path, LinkToFile=False, SaveWithDocument=True,
                                      Left=left_pos, Top=top_pos,
                                      Width=pic_width, Height=pic_height)
        shape.WrapFormat.Type = 3
        shape.Line.Weight = 1
        shape.Line.ForeColor.RGB = gray_color

        left_pos += pic_width

        if (i + 1) % pics_per_row == 0:
            left_pos = left_start
            top_pos += pic_height

log_action(f"Word file for Package {package} created successfully.\n")
print(f"Word file for Package {package} created successfully.")
