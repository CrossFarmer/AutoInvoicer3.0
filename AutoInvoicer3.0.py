import sys
import cv2
import pytesseract
import os
import shutil
from pdf2image import convert_from_path
import re
import datetime
from PIL import Image, ImageFilter
import tkinter as tk
from tkinter import filedialog
import datetime
import threading
import ctypes

# You will need to provide the path to the tesseract executable in your system
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

class PDFProcessor:
        def __init__(self):
            self.INV_DATE_LOCS = [(1224, 224, 1405, 361), (1222, 244, 1407, 386), (1199, 259, 1421, 357), (1201, 243, 1386, 385), (1200, 222, 1385, 366), (1203, 200, 1388, 342), (1222, 196, 1407, 338), (1241, 197, 1426, 339), (1243, 221, 1428, 363), (1243, 242, 1428, 384), (1227, 269, 1401, 335), (1227, 247, 1401, 313), (1193, 267, 1367, 333), (1226, 285, 1400, 351), (1253, 269, 1427, 335), (1216, 262, 1413, 341), (1263, 289, 1369, 323), (1248, 284, 1354, 318), (1245, 295, 1351, 329), (1263, 298, 1369, 332), (1282, 298, 1388, 332), (1290, 289, 1396, 323), (1283, 278, 1389, 312)]
            self.INV_NUM_LOCS = [(1395, 224, 1603, 365), (1393, 241, 1601, 382), (1374, 239, 1582, 380), (1364, 224, 1572, 365), (1362, 209, 1570, 350), (1370, 200, 1578, 341), (1393, 196, 1601, 337), (1429, 204, 1637, 345), (1439, 242, 1647, 383), (1399, 214, 1600, 345), (1429, 280, 1574, 345), (1410, 266, 1555, 331), (1444, 265, 1589, 330), (1411, 288, 1556, 353), (1445, 287, 1590, 352), (1395, 273, 1593, 337), (1402, 279, 1595, 339)]
            self.CUST_LOCS = [(92, 607, 608, 920), (93, 669, 609, 982), (31, 665, 547, 978), (1, 609, 517, 922), (2, 556, 518, 869), (90, 546, 606, 859), (103, 590, 619, 903), (104, 613, 620, 926), (104, 656, 620, 969), (30, 638, 318, 818), (2, 641, 290, 821), (45, 578, 333, 758), (90, 661, 378, 841), (99, 675, 284, 787)]
            self.company_names = self.read_company_names()
            self.bundle_dir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
            self.poppler_path = os.path.join(self.bundle_dir, 'poppler/bin')

            self.inv_date_pattern = re.compile(r'(\d{1,2}[\/\\]\d{1,2}[\/\\]\d{4})')
            self.inv_num_pattern = re.compile(r'(A\d{6})')

        def read_company_names(self):
            try:
                with open('company_names.txt', 'r') as file:
                    # This will read each line of the file into a list element, stripping the newline characters
                    company_names = [line.strip() for line in file.readlines()]
                return company_names
            except FileNotFoundError:
                print("The company_names.txt file was not found in the same directory as the script.")
                return []

        def extract_text(self, image, bbox):
            image.convert('L')
            roi = image.crop(bbox)
            text = pytesseract.image_to_string(roi)
            return text
        
        def extract_text_2(self, image, bbox):
            image.convert('L')
            image = image.point(lambda x: x > 233 and 255)
            #image = image.filter(ImageFilter.SHARPEN)
            #image = image.filter(ImageFilter.SMOOTH_MORE)
            #image = image.filter(ImageFilter.BoxBlur(radius=1))
            image = image.filter(ImageFilter.GaussianBlur(radius=1))
            roi = image.crop(bbox)
            text = pytesseract.image_to_string(roi)
            return text
        
        def extract_invoice_number(self):
            inv_num = 'UNKNOWN'
            for inv_num_loc in self.INV_NUM_LOCS:
                for extract_method in [self.extract_text, self.extract_text_2]:
                    inv_num_text = extract_method(self.image, inv_num_loc)
                    # Try to match the extracted text with the pattern
                    match = self.inv_num_pattern.search(inv_num_text)
                    if match:
                        inv_num = match.group()
                        break  # If a match is found, break from the inner loop
                if inv_num != 'UNKNOWN':
                    break  # If a valid invoice number was found, break from the outer loop as well

            if inv_num == 'UNKNOWN':
                print('Invoice number not found.')

            return inv_num


        def extract_invoice_date(self):
            inv_date = 'UNKNOWN'
            for inv_date_loc in self.INV_DATE_LOCS:
                for extract_method in [self.extract_text, self.extract_text_2]:
                    inv_date_text = extract_method(self.image, inv_date_loc)
                    inv_date_match = self.inv_date_pattern.search(inv_date_text)
                    if inv_date_match:
                        try:
                            # Try to parse the matched date string as a date.
                            parsed_date = datetime.datetime.strptime(inv_date_match.group(), '%m/%d/%Y')
                            # If the date is parsed successfully, format it as desired and break from the inner loop.
                            inv_date = parsed_date.strftime('%m.%d.%Y')
                            break
                        except ValueError:
                            # This will be triggered if the date is unrealistic,
                            # in which case we ignore it and move on to the next extraction method.
                            continue
                # If a valid date was found, break from the outer loop as well.
                if inv_date != 'UNKNOWN':
                    break

            if inv_date == 'UNKNOWN':
                print('Invoice date not found.')

            return inv_date


        def extract_company_name(self):
            company = 'UNKNOWN'
            for cust_loc in self.CUST_LOCS:
                for extract_method in [self.extract_text, self.extract_text_2]:
                    cust_text = extract_method(self.image, cust_loc)
                    for name in self.company_names:
                        if name in cust_text:
                            company = name
                            break  # If a match is found, break from the innermost loop
                    if company != 'UNKNOWN':
                        break  # If a valid company name was found, break from the middle loop
                if company != 'UNKNOWN':
                    break  # If a valid company name was found, break from the outermost loop

            if company == 'UNKNOWN':
                print('Company name not found.')
            
            return company


        def process_pdf(self, file_path, destination):
            images = convert_from_path(file_path, poppler_path=self.poppler_path)
            # Process only the first page
            image = images[0]

            self.image = image.resize((1700, 2200))

            inv_num = self.extract_invoice_number()
            inv_date = self.extract_invoice_date()
            company = self.extract_company_name()

            print(f'Final extracted data: inv_date: "{inv_date}", inv_num: "{inv_num}", company: "{company}"')
            base_filename = f"{inv_date}_{inv_num}"
            new_filename = base_filename + ".pdf"
            print(f'Created filename: "{new_filename}"')

            if 'UNKNOWN' in new_filename and 'UNKNOWN' in company:
                failed_dir = os.path.join(destination, 'FAILED')
                os.makedirs(failed_dir, exist_ok=True)
                shutil.copy(file_path, failed_dir)
            else:
                destination_dir = os.path.join(destination, company)
                os.makedirs(destination_dir, exist_ok=True)

                destination_file_path = os.path.join(destination_dir, new_filename)

                copy_num = 1
                while os.path.exists(destination_file_path):
                    new_filename = f"{base_filename}_{copy_num}.pdf"
                    destination_file_path = os.path.join(destination_dir, new_filename)
                    copy_num += 1

                shutil.move(file_path, destination_file_path)

            # Delete the original file after all processing is done
            if os.path.exists(file_path):
                os.remove(file_path)



        def process_directory(self, source, destination):
            for file in os.listdir(source):
                if file.endswith(".pdf"):
                    self.process_pdf(os.path.join(source,file), destination)
            print("ALL INVOICES RENAMED")


class StdoutRedirector(object):
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, string):
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)  # Scrolls to the end

    def flush(self):
        pass  # This might be necessary if you're using Python 3



class MainWindow:
    def __init__(self, master):
        self.master = master
        self.master.title("File Processor")
        self.master.geometry("900x650")

        self.source_dir = ""
        self.destination_dir = ""

        self.buttons_frame = tk.Frame(self.master)
        self.buttons_frame.place(relx=0.50, rely=0.89, relheight=0.15, relwidth=0.95, anchor='center')

        self.title_frame = tk.Frame(self.master, bd=5, relief="groove")
        self.title_frame.place(relx=0.5, rely=0.05, relheight=0.10, relwidth=0.95, anchor='center')

        self.notice_frame = tk.Frame(self.master, bd=5, relief="groove")
        self.notice_frame.place(relx=0.5, rely=0.47, relheight=0.70, relwidth=0.95, anchor='center')

        self.title_label_title = tk.Label(self.title_frame, text="AUTO INVOICER 3.0", font=('Helvetica', 14, 'bold'))
        self.title_label_title.place(relx=0.12, rely=0.50, relheight=0.95, relwidth=0.23, anchor='center')

        self.title_label_notes = tk.Label(self.title_frame, text="Make sure to make a copy of your input folder, as it's contents are deleted after renaming process completes!", font=('Times New Roman', 10, 'italic bold'))
        self.title_label_notes.place(relx=0.63, rely=0.50, relheight=0.95, relwidth=0.73, anchor='center')

        self.output_text = tk.Text(self.notice_frame, wrap='word')
        self.output_text.pack(fill='both', expand=True)

        self.source_button = tk.Button(self.buttons_frame, text="INPUT FOLDER", command=self.select_source, font=('Arial', 12))
        self.source_button.place(relx=0.17, rely=0.55, relheight=0.60, relwidth=0.20, anchor='center')

        self.destination_button = tk.Button(self.buttons_frame, text="OUTPUT FOLDER", command=self.select_destination, font=('Arial', 12))
        self.destination_button.place(relx=0.41, rely=0.55, relheight=0.60, relwidth=0.20, anchor='center')

        self.run_button = tk.Button(self.buttons_frame, text="START PROCESS", command=self.run,  font=('Arial', 16, 'bold italic'), state=tk.DISABLED)
        self.run_button.place(relx=0.80, rely=0.55, relheight=0.65, relwidth=0.25, anchor='center')

        self.processor = PDFProcessor()
        sys.stdout = StdoutRedirector(self.output_text)

    def update_button_state(self):
        if self.source_dir and self.destination_dir:
            self.run_button['state'] = tk.NORMAL
        else:
            self.run_button['state'] = tk.DISABLED

    def select_source(self):
        self.source_dir = filedialog.askdirectory()
        if self.source_dir:
            self.source_button['text'] = "✓ INPUT FOLDER"
            self.source_button['fg'] = "dark green"
            self.source_button['font'] = ('Arial', 12, 'bold')
        self.update_button_state()
        print(f"Selected source directory: {self.source_dir}")

    def select_destination(self):
        self.destination_dir = filedialog.askdirectory()
        if self.destination_dir:
            self.destination_button['text'] = "✓ OUTPUT FOLDER"
            self.destination_button['fg'] = "dark green"
            self.destination_button['font'] = ('Arial', 12, 'bold')
        self.update_button_state()
        print(f"Selected destination directory: {self.destination_dir}")

    def run(self):
        print("PROCESSING INVOICE FILES")
        if not self.source_dir or not self.destination_dir:
            print("Please select both source and destination directories.")
            return
        # Disable the button to prevent multiple clicks
        self.run_button['state'] = tk.DISABLED
        # Start the long-running process in a new thread
        process_thread = threading.Thread(target=self.start_process, daemon=True)
        process_thread.start()

    def start_process(self):
        # Call the process_directory method of your PDFProcessor instance
        # This function will run in a separate thread
        self.processor.process_directory(self.source_dir, self.destination_dir)
        # Re-enable the button after the process is complete
        self.run_button['state'] = tk.NORMAL


def minimize_console():
    whnd = ctypes.windll.kernel32.GetConsoleWindow()
    if whnd != 0:
        ctypes.windll.user32.ShowWindow(whnd, 6)
        ctypes.windll.user32.SetWindowPos(whnd, 0, -100, -100, 0, 0, 0x0001)

minimize_console()
root = tk.Tk()
app = MainWindow(root)
root.mainloop()
