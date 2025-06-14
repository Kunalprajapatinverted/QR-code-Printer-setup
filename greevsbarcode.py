import win32print
import win32ui
import qrcode
import win32com.client
from io import BytesIO
import sys
from PIL import Image, ImageWin

def get_usb_devices():  # Printer connection check
    wmi = win32com.client.GetObject('winmgmts:')
    usb_devices = wmi.ExecQuery('SELECT * FROM Win32_USBHub')
    printer_connected = False
    for device in usb_devices:
        if "VID_1203&PID_0272" in device.DeviceID:
            printer_connected = True

    if printer_connected:
        print("TSC TE244 printer is connected.")
    else:
        sys.exit("TSC TE244 printer is not connected.\nPlease check your printer connection and try again.")

def print_key_value_data(data_dict):  # Text appears side-by-side with QR code
    printer_name = win32print.GetDefaultPrinter()
    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(printer_name)
    dc.StartDoc("Key-Value Print")
    dc.StartPage()

    font_name = "Arial"
    font_size = 38 # Text Size
    line_spacing = 10
    margin_left = 120
    margin_top = 50

    # Set up font
    font = win32ui.CreateFont({
        "name": font_name,
        "height": font_size,
        "weight": 500 # text bold
    })
    dc.SelectObject(font)

    y_position = margin_top
    final_text = "PDI OK"

    # Print key-value pairs except 'PDI OK'
    for key, value in data_dict.items():
        if key.lower() == "pdi ok":
            continue
        line = f"{key} : {value}"
        dc.TextOut(margin_left, y_position, line)
        y_position += font_size + line_spacing

    # Extract battery code from either key or value
    battery_code = None
    for k, v in data_dict.items():
        for candidate in [k, v]:
            if (
                isinstance(candidate, str) and
                candidate.startswith("IN") and
                len(candidate) >= 16 and
                candidate[-1].isdigit()
            ):
                battery_code = candidate
                break
        if battery_code:
            break

    qr_data = battery_code if battery_code else "UNKNOWN"

    # Generate QR code
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=4, #QRCode size 
        border=0,
    )
    qr.add_data(qr_data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")

    # Convert image to in-memory BMP and load
    buffer = BytesIO()
    img.save(buffer, format="BMP")
    buffer.seek(0)
    bmp = Image.open(buffer)
    dib = ImageWin.Dib(bmp)

    qr_width, qr_height = bmp.size
    spacing = 40  # Space between text and QR
    block_y = y_position + 60
    text_x = margin_left
    text_y = block_y + (qr_height - font_size) // 2
    qr_x = text_x + dc.GetTextExtent(final_text)[0] + spacing
    qr_y = block_y

    # Draw final text and QR code side-by-side
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
        try:
            print_key_value_data(data_dict)
            print("Label printed successfully.")
        except Exception as e:
            print(f"Error while printing: {e}")
    else:
        print("No data entered. Nothing to print.")
