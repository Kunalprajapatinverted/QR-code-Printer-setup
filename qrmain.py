import win32print
import win32ui
import qrcode
import win32com.client
import sys
from PIL import Image, ImageWin
'''
TSC TE244 when direct connect with system
USB Printing Support when connect via extended usb support cabel
Device Name: USB Printing Support
Device ID: USB\\VID_1203&PID_0272\000001
PNP Device ID: USB\\VID_1203&PID_0272\000001 
'''
def get_usb_devices(): #printer connection check
    wmi = win32com.client.GetObject('winmgmts:')    # Initialize the WMI client
    usb_devices = wmi.ExecQuery('SELECT * FROM Win32_USBHub') # Query connected USB devices
    printer_connected = False# Initialize a flag to check for the printer
    for device in usb_devices:# Print details of each USB device
        # Check if this is the TSC te244 printer
        if (device.DeviceID == "USB\\VID_1203&PID_0272\\000001"):
            ab = device.name
            printer_connected = True

    # Final check to see if the printer was found
    if printer_connected:
        print("TSC te244 printer is connected.")
    else:
        sys.exit("TSC te244 printer is not connected. \nPlease check your printer connection and try again")

def print_qr_code(text=None, additional_text=None, second_text=None, second_additional_text=None):
    # Get the default printer
    printer_name = win32print.GetDefaultPrinter()

    # Create a device context object
    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(printer_name)

    # Start the document
    dc.StartDoc("Print Job")
    dc.StartPage()

    # Set the font for the text
    font = "Arial"
    max_font_size = 30
    # min_font_size = 10
    max_text_width = 400             
    max_text_height = 250 

    # Function to calculate font size
    def calculate_font_size(text):
        font_size = max_font_size
        dc.SelectObject(win32ui.CreateFont({
            "name": font,
            "height": font_size,
            "weight": 500
        }))
        text_width, text_height = dc.GetTextExtent(text) #adjust the size of the text 
        while text_width > max_text_width or text_height > max_text_height:
            font_size -= 1
            dc.SelectObject(win32ui.CreateFont({
                "name": font,
                "height": font_size,
                "weight": 500
            }))
            text_width, text_height = dc.GetTextExtent(text)  #recalculates the text width and height based on the new font size
        return text_width, text_height, font_size

    # Draw the first text and additional text
    if text:
        text_width, text_height, font_size = calculate_font_size(text)
        x_coord = (max_text_width - text_width) // 2 + 7
        # print(x_coord)
        dc.TextOut(x_coord, 7, text)

        if additional_text:
            additional_text_width, additional_text_height = dc.GetTextExtent(additional_text)
            available_space = max_text_width
            x_coord_additional = (available_space - additional_text_width) // 2 + 10
            # print(x_coord_additional)                 
            dc.TextOut(x_coord_additional, text_height + 7, additional_text)

    # Draw the second text and second additional text in front of the first text 
    if second_text:
        second_text_width, second_text_height, font_size = calculate_font_size(second_text)
        
        # Adjust the margin between the two texts based on their lengths
        if len(text) == 12 and len(second_text) == 12:
            margin = 170
        elif 12 < len(text) <= 16 and 12 < len(second_text) <= 16:
            margin = 130
        elif 16 < len(text) <= 20 and 16 < len(second_text) <= 20:
            margin = 80
        elif 20 < len(text) <= 24 and 20 < len(second_text) <= 24:
            margin = 10
        else:
            margin = 50  # Default margin
        
        x_coord_second = x_coord + text_width + margin  # Position it next to the first text
        dc.TextOut(x_coord_second, 10, second_text)

        if second_additional_text:  
            second_additional_text_width, second_additional_text_height = dc.GetTextExtent(second_additional_text)
            # Center the second_additional_text relative to second_text
            x_coord_second_additional = x_coord_second + (second_text_width - second_additional_text_width) // 2
            dc.TextOut(x_coord_second_additional, text_height + 10, second_additional_text)

    # Generate the QR code for the first text
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=4,
        border=0,
    )
    qr.add_data(text)
    qr.make(fit=True)

    # Create an image from the QR code for the first text
    img = qr.make_image(fill_color="black", back_color="white")
    img.save("firstQR.png")

    # Load the image into a Windows bitmap
    bmp = Image.open("firstQR.png")
    dib = ImageWin.Dib(bmp)
    dib_width, dib_height = dib.size

    # Calculate the position of the QR code below the first set of text
    qr_x = 150
    qr_y = text_height + (additional_text_height if additional_text else 0) + 20
    dib.draw(dc.GetHandleOutput(), (qr_x, qr_y, qr_x + dib_width, qr_y + dib_height))

    # Generate the QR code for the second text
    qr_second = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=4,
        border=0,
    )
    print(second_text)
    qr_second.add_data(second_text)
    qr_second.make(fit=True)

    # Create an image from the QR code for the second text
    img_second = qr_second.make_image(fill_color="black", back_color="white")
    img_second.save("secondQR.png")

    # Load the image into a Windows bitmap
    bmp_second = Image.open("secondQR.png")
    dib_second = ImageWin.Dib(bmp_second)
    dib_second_width, dib_second_height = dib_second.size

    # Calculate the position of the QR code for the second text
    qr_x_second = qr_x + dib_width + 350  # Position it next to the first QR code
    qr_y_second = qr_y  # Align it with the first QR code
    dib_second.draw(dc.GetHandleOutput(), (qr_x_second, qr_y_second, qr_x_second + dib_second_width, qr_y_second + dib_second_height))
    
    # End the page and document
    dc.EndPage()
    dc.EndDoc()
    dc.DeleteDC()

# def view_print_queue(): #printer Queue view
#     printer_name = win32print.GetDefaultPrinter()   # Get the default printer
#     hPrinter = win32print.OpenPrinter(printer_name) # Open the printer
#     try:
#         # Get the print jobs
#         print_jobs = win32print.EnumJobs(hPrinter, 0, 999, 1)  # Get all jobs
#         if not print_jobs:
#             print("No print jobs in the queue.")
#             return
#         print(f"Print jobs for printer: {printer_name}")
#         for i, job in enumerate(print_jobs):
#             print(f"Job {i+1}: ID: {job.get('JobId', 'Unknown')}, Document: {job.get('Document', 'Unknown')}, Status: {job.get('Status', 'Unknown')}") 
#     finally:
#         # Close the printer handle
#         win32print.ClosePrinter(hPrinter)

def view_print_queue():
    printer_name = win32print.GetDefaultPrinter()  # Get the default printer
    hPrinter = win32print.OpenPrinter(printer_name)  # Open the printer
    try:
        # Get the print jobs
        print_jobs = win32print.EnumJobs(hPrinter, 0, 999, 1)  # Get all jobs
        if not print_jobs:
            print("No print jobs in the queue.")
            return
        
        print(f"Print jobs for printer: {printer_name}")
        for i, job in enumerate(print_jobs):
            print(f"Job {i + 1}:")
            for key, value in job.items():
                print(f"  {key}: {value}")
    finally:
        # Close the printer handle
        win32print.ClosePrinter(hPrinter)

if __name__ == "__main__":
    view_print_queue()
    print(view_print_queue)
    get_usb_devices()
    while True:
        # Get all inputs from the user as strings
        text = input("Enter the first text (required): ")
        additional_text = input("Enter additional info for the first text (optional): ")
        second_text = input("Enter the second text (optional): ")
        second_additional_text = input("Enter additional info for the second text (optional): ")

        # Check if required fields are filled
        if text:
            # Check length of text
            if not (12 <= len(text) <= 24):
                print("Please input text length in the range between 12 to 24 characters.")
                continue  # Go back to the beginning of the loop

            if not second_text: # If second text is not provided, use the first text
                second_text = text
            if not (12 <= len(second_text) <= 24): # Check length of second_text
                print("Please input second text length in the range between 12 to 24 characters.")
                continue  # Go back to the beginning of the loop
            
            # If second additional text is not provided, use the first additional text
            if not second_additional_text:
                second_additional_text = additional_text

            print_qr_code(text, additional_text, second_text, second_additional_text)
            print("QR codes printed successfully.")
            break  # Exit the loop if everything is valid
        else:
            print("The first text is required. Please try again.")