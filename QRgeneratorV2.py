import os
import openpyxl
import qrcode
import re
import cv2
import numpy as np
from PIL import Image

def sanitize_filename(filename):
    """
    Sanitize a filename by removing invalid characters.
    """
    return re.sub(r'[<>:"/\\|?*]', '', filename)

def generate_qr_code(data):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(data)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white")
    return qr_img

def find_green_square(image):
    hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)

    lower_green = np.array([35, 50, 50])
    upper_green = np.array([85, 255, 255])
    mask = cv2.inRange(hsv, lower_green, upper_green)
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    contours = sorted(contours, key=cv2.contourArea, reverse=True)

    for contour in contours:
        peri = cv2.arcLength(contour, True)
        approx = cv2.approxPolyDP(contour, 0.02 * peri, True)
        if len(approx) == 4:
            return approx.reshape((-1, 2))

    return None

workbook = openpyxl.load_workbook('data.xlsx')
sheet = workbook.active

#Create folders if they don't exist
os.makedirs('E-Tickets', exist_ok=True)
os.makedirs('info', exist_ok=True)

row_count = 1
for row in sheet.iter_rows(min_row=2, values_only=True):
    data = '\n'.join(str(cell) for cell in row if cell is not None)
    qr_code = generate_qr_code(data)

    #Convert the QR code image to a NumPy array
    qr_code_array = np.array(qr_code.convert('RGB')) 

    #Template loading
    template_path = 'ticket_template.png'
    template_img = cv2.imread(template_path)

    #Replace QR instead of green square on the ticket
    green_square = find_green_square(template_img)
    if green_square is not None:
        
        x, y, w, h = cv2.boundingRect(green_square)
        qr_code_resized = cv2.resize(qr_code_array, (w, h))
        template_img[y:y+h, x:x+w] = qr_code_resized
        ticket_id = re.findall(r'Ticket ID:\s*(\w+)', data)
        if ticket_id:
            ticket_id = ticket_id[0]

    #Generate unique filenames
    filename = f'stdinfo{row_count}'
    output_path = os.path.join('E-Tickets', f'{filename}.png')
    cv2.imwrite(output_path, template_img)
    print(f'Generated E-ticket for row {row_count} as {output_path}')

    text_filename = os.path.join('info', f'{filename}.txt')
    with open(text_filename, 'w', encoding='utf-8') as text_file:
        text_file.write(data)
    print(f'Saved row {row_count} data to {text_filename}')

    row_count += 1


#แก้ empty roll เพิ่มเติม

