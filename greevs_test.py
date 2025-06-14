import win32print
import win32ui
import qrcode
import win32com.client
from io import BytesIO
import sys
from PIL import Image, ImageWin

def get_usb_devices():  # printer connection check
    wmi = win32com.client.GetObject('winmgmts:')
    usb_devices = wmi.ExecQuery('SELECT * FROM Win32_USBHub')
    printer_connected = False
    for device in usb_devices:
        if (device.DeviceID == "USB\\VID_1203&PID_0272\\000001"):
            printer_connected = True

    if printer_connected:
        print("TSC TE244 printer is connected.")
    else:
        sys.exit("TSC TE244 printer is not connected.\nPlease check your printer connection and try again.")

def print_key_value_data(data_dict):
    printer_name = win32print.GetDefaultPrinter()
    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(printer_name)
    dc.StartDoc("Key-Value Print")
    dc.StartPage()

    font_name = "Arial"
    font_size = 38
    line_spacing = 10
    margin_left = 120 #50
    margin_top = 50

    # Set up font
    font = win32ui.CreateFont({
        "name": font_name,
        "height": font_size,
        "weight": 500
    })
    dc.SelectObject(font)

    # Calculate alignment offset
    max_key_width = max(dc.GetTextExtent(key + " : ")[0] for key in data_dict.keys())

    y_position = margin_top
    qr_data = ""

    for key, value in data_dict.items():
        line = f"{key} : {value}"
        dc.TextOut(margin_left, y_position, line)
        y_position += font_size + line_spacing
        qr_data += f"{key}:{value};"

    # Generate QR from full data string
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=4,
        border=0,
    )
    qr.add_data(qr_data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img.save("kv_qr.png")

    # Load and print QR
    bmp = Image.open("kv_qr.png")
    dib = ImageWin.Dib(bmp)
    qr_x = margin_left + 100
    qr_y = y_position + 20
    dib.draw(dc.GetHandleOutput(), (qr_x, qr_y, qr_x + bmp.size[0], qr_y + bmp.size[1]))

    dc.EndPage()
    dc.EndDoc()
    dc.DeleteDC() #QR code after the text 

def view_print_queue():
    printer_name = win32print.GetDefaultPrinter()
    hPrinter = win32print.OpenPrinter(printer_name)
    try:
        print_jobs = win32print.EnumJobs(hPrinter, 0, 999, 1)
        if not print_jobs:
            print("No print jobs in the queue.")
            return

        print(f"Print jobs for printer: {printer_name}")
        for i, job in enumerate(print_jobs):
            print(f"Job {i + 1}:")
            for key, value in job.items():
                print(f"  {key}: {value}")
    finally:
        win32print.ClosePrinter(hPrinter)

if __name__ == "__main__":
    view_print_queue()
    get_usb_devices()

    data_dict = {}
    print("Enter the information for the label. Type 'done' as key when finished.")
    while True:
        key = input("Enter key (e.g. First Name): ")
        if key.lower() == "done":
            break
        value = input(f"Enter value for '{key}': ")
        data_dict[key] = value

    if data_dict:
        print_key_value_data(data_dict)
        print("Label printed successfully.")
    else:
        print("No data entered. Nothing to print.")

# *******************************************************************************

import win32print
import win32ui
import qrcode
import win32com.client
from io import BytesIO
import sys
from PIL import Image, ImageWin

def get_usb_devices():  # printer connection check
    wmi = win32com.client.GetObject('winmgmts:')
    usb_devices = wmi.ExecQuery('SELECT * FROM Win32_USBHub')
    printer_connected = False
    for device in usb_devices:
        if (device.DeviceID == "USB\\VID_1203&PID_0272\\000001"):
            printer_connected = True

    if printer_connected:
        print("TSC TE244 printer is connected.")
    else:
        sys.exit("TSC TE244 printer is not connected.\nPlease check your printer connection and try again.")

def print_key_value_data(data_dict):
    printer_name = win32print.GetDefaultPrinter()
    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(printer_name)
    dc.StartDoc("Key-Value Print")
    dc.StartPage()

    font_name = "Arial"
    font_size = 38
    line_spacing = 10
    margin_left = 170
    margin_top = 50

    # Set up font
    font = win32ui.CreateFont({
        "name": font_name,
        "height": font_size,
        "weight": 500
    })
    dc.SelectObject(font)

    y_position = margin_top
    qr_data = ""

    # Print key-value pairs
    for key, value in data_dict.items():
        if key.lower() == "pdi ok":
            continue  # Skip this; print it separately
        line = f"{key} : {value}"
        dc.TextOut(margin_left, y_position, line)
        y_position += font_size + line_spacing
        qr_data += f"{key}:{value};"

    # Generate QR from full data string
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=4,
        border=0,
    )
    qr.add_data(qr_data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img.save("kv_qr.png")

    # Load QR image
    bmp = Image.open("kv_qr.png")
    dib = ImageWin.Dib(bmp)

    qr_width, qr_height = bmp.size
    qr_x = margin_left + 100
    qr_y = y_position + 60  # Space between text and QR

    # Draw final message above QR, centered
    final_text = "PDI OK"
    final_text_width = dc.GetTextExtent(final_text)[0]
    text_x = qr_x + (qr_width - final_text_width) // 2
    text_y = qr_y - font_size - 10  # Position just above QR code
    dc.TextOut(text_x, text_y, final_text)

    # Draw QR Code
    dib.draw(dc.GetHandleOutput(), (qr_x, qr_y, qr_x + qr_width, qr_y + qr_height))

    dc.EndPage()
    dc.EndDoc()
    dc.DeleteDC() #QR code just down of the text and text is in centre

def view_print_queue():
    printer_name = win32print.GetDefaultPrinter()
    hPrinter = win32print.OpenPrinter(printer_name)
    try:
        print_jobs = win32print.EnumJobs(hPrinter, 0, 999, 1)
        if not print_jobs:
            print("No print jobs in the queue.")
            return

        print(f"Print jobs for printer: {printer_name}")
        for i, job in enumerate(print_jobs):
            print(f"Job {i + 1}:")
            for key, value in job.items():
                print(f"  {key}: {value}")
    finally:
        win32print.ClosePrinter(hPrinter)

if __name__ == "__main__":
    view_print_queue()
    get_usb_devices()

    data_dict = {}
    print("Enter the information for the label. Type 'done' as key when finished.")
    while True:
        key = input("Enter key (e.g. First Name): ")
        if key.lower() == "done":
            break
        value = input(f"Enter value for '{key}': ")
        data_dict[key] = value

    if data_dict:
        print_key_value_data(data_dict)
        print("Label printed successfully.")
    else:
        print("No data entered. Nothing to print.")

# *******************************************************************************

import win32print
import win32ui
import qrcode
import win32com.client
from io import BytesIO
import sys
from PIL import Image, ImageWin

def get_usb_devices():  # printer connection check
    wmi = win32com.client.GetObject('winmgmts:')
    usb_devices = wmi.ExecQuery('SELECT * FROM Win32_USBHub')
    printer_connected = False
    for device in usb_devices:
        if (device.DeviceID == "USB\\VID_1203&PID_0272\\000001"):
            printer_connected = True

    if printer_connected:
        print("TSC TE244 printer is connected.")
    else:
        sys.exit("TSC TE244 printer is not connected.\nPlease check your printer connection and try again.")

def print_key_value_data(data_dict):  #text appears side-by-side with the QR code
    printer_name = win32print.GetDefaultPrinter()
    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(printer_name)
    dc.StartDoc("Key-Value Print")
    dc.StartPage()

    font_name = "Arial"
    font_size = 38
    line_spacing = 10
    margin_left = 120
    margin_top = 50

    # Set up font
    font = win32ui.CreateFont({
        "name": font_name,
        "height": font_size,
        "weight": 500
    })
    dc.SelectObject(font)

    y_position = margin_top
    qr_data = ""
    final_text = "PDI OK"

    # Print key-value pairs except 'PDI OK'
    for key, value in data_dict.items():
        if key.lower() == "pdi ok":
            continue
        line = f"{key} : {value}"
        dc.TextOut(margin_left, y_position, line)
        y_position += font_size + line_spacing
        qr_data += f"{key}:{value};"

    # Generate QR code
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=4,
        border=0,
    )
    qr.add_data(qr_data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img.save("kv_qr.png")

    # Load QR image
    bmp = Image.open("kv_qr.png")
    dib = ImageWin.Dib(bmp)

    qr_width, qr_height = bmp.size
    spacing = 40  # Space between text and QR
    total_block_height = max(font_size, qr_height)

    # Position for side-by-side layout
    block_y = y_position + 60
    text_x = margin_left
    text_y = block_y + (qr_height - font_size) // 2

    qr_x = text_x + dc.GetTextExtent(final_text)[0] + spacing
    qr_y = block_y

    # Draw text and QR side-by-side
    dc.TextOut(text_x, text_y, final_text)
    dib.draw(dc.GetHandleOutput(), (qr_x, qr_y, qr_x + qr_width, qr_y + qr_height))

    dc.EndPage()
    dc.EndDoc()
    dc.DeleteDC()


def view_print_queue():
    printer_name = win32print.GetDefaultPrinter()
    hPrinter = win32print.OpenPrinter(printer_name)
    try:
        print_jobs = win32print.EnumJobs(hPrinter, 0, 999, 1)
        if not print_jobs:
            print("No print jobs in the queue.")
            return

        print(f"Print jobs for printer: {printer_name}")
        for i, job in enumerate(print_jobs):
            print(f"Job {i + 1}:")
            for key, value in job.items():
                print(f"  {key}: {value}")
    finally:
        win32print.ClosePrinter(hPrinter)

if __name__ == "__main__":
    view_print_queue()
    get_usb_devices()

    data_dict = {}
    print("Enter the information for the label. Type 'done' as key when finished.")
    while True:
        key = input("Enter key (e.g. First Name): ")
        if key.lower() == "done":
            break
        value = input(f"Enter value for '{key}': ")
        data_dict[key] = value

    if data_dict:
        print_key_value_data(data_dict)
        print("Label printed successfully.")
    else:
        print("No data entered. Nothing to print.")
