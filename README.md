<h1> QRCode Generater for user information</h1>

This Python script generates electronic tickets (E-Tickets) using data from an Excel file (`data.xlsx`). 
Each row in the Excel file represents a ticket, and the script creates a corresponding E-Ticket image and saves the ticket information to a text file.
One problem that haven't fix yet is the program will generate empty roll -- doesn't stop when find empty row (maybe that row seem empty but not null)

<h2> Use Case</h2>
This program is used for the RakNong66 Event at Mahidol University. The event aims to welcome freshmen to Mahidol and includes activities such as ice-breaking, making new friends, providing information about Mahidol, and a free concert.

<h2> Dependencies </h2>
- `os`: For interacting with the operating system.
- `openpyxl`: For reading Excel files.
- `qrcode`: For generating QR codes.
- `re`: For regular expressions.
- `cv2` (OpenCV): For image processing.
- `numpy`: For numerical computing.
- `PIL`: For image processing.

<h2> Installation </h2>
1. Install the required dependencies: "pip install openpyxl qrcode opencv-python-headless numpy pillow"
2. Place the `data.xlsx` file containing ticket information in the same directory as the script.

<h2> Usage </h2>

1. Ensure the `data.xlsx` file is correctly formatted, with each row representing a ticket and each column containing relevant ticket information.
2. The script will generate E-Ticket images in the `E-Tickets` folder and corresponding text files in the `info` folder.

<h2> Note </h2>
- Ensure that the ticket template (`ticket_template.png`) contains a green square placeholder where the QR code will be inserted.
- The color code of green square (RGB): #1CFF00 





