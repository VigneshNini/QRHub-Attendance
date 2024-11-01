import cv2
from pyzbar.pyzbar import decode
import openpyxl
from openpyxl.styles import Font
from datetime import datetime
import tkinter as tk
import tkinter.font as tkFont
import os
import threading

class QRScannerApp:
    def __init__(self):
        self.cap = cv2.VideoCapture(0)
        self.recorded_qr_codes = []

        self.root = tk.Tk()
        self.root.title("QR Code Attendance Scanner")
        self.font = tkFont.Font(family="Helvetica", size=14, weight="bold")

        self.setup_gui()
        self.start_video_stream()

    def setup_gui(self):
        self.excel_file = 'Attendance_Log.xlsx'
        try:
            self.workbook = openpyxl.load_workbook(self.excel_file)
        except FileNotFoundError:
            self.workbook = openpyxl.Workbook()

        self.sheet = self.workbook.active
        headings = ["Student ID", "Name", "Time", "Date"]
        self.sheet.append(headings)
        for cell in self.sheet["1:1"]:
            cell.font = Font(bold=True)

        open_excel_button = tk.Button(self.root, text="Open Excel Folder", command=self.open_excel_folder, font=self.font)
        open_excel_button.pack(pady=10)

        stop_button = tk.Button(self.root, text="Stop Scanner", command=self.stop_scanner, font=self.font)
        stop_button.pack(pady=10)

        self.video_frame = tk.Frame(self.root)
        self.video_frame.pack()
        self.video_label = tk.Label(self.video_frame)
        self.video_label.pack()

        self.date_label = tk.Label(self.root, text="Date: ", font=("Helvetica", 14, "bold"))
        self.date_label.pack(pady=5)

        self.time_label = tk.Label(self.root, text="Time: ", font=("Helvetica", 14, "bold"))
        self.time_label.pack(pady=5)

        self.success_message_label = tk.Label(self.root, text="", font=("Helvetica", 24, "bold"), fg="blue")

    def open_excel_folder(self):
        folder_path = os.path.dirname(os.path.abspath(self.excel_file))
        os.startfile(folder_path)

    def register_attendance(self, qr_code_content):
        student_id = ""
        student_name = ""
        parts = qr_code_content.split("\n")
        for part in parts:
            if part.startswith("Student ID: "):
                student_id = part.replace("Student ID: ", "")
            elif part.startswith("Name: "):
                student_name = part.replace("Name: ", "")

        if not student_id or not student_name:
            return  # Invalid QR code content

        current_datetime = datetime.now()
        current_time = current_datetime.strftime("%I:%M %p")
        current_date = current_datetime.strftime("%d-%m-%Y")

        self.sheet.append([student_id, student_name, current_time, current_date])
        self.workbook.save(self.excel_file)
        self.recorded_qr_codes.append(qr_code_content)

        top_message_label = tk.Label(self.root, text=f"Attendance Recorded for {student_name} (ID: {student_id})",
                                font=("Helvetica", 20, "bold"), fg="green")
        top_message_label.pack(pady=10)

        self.date_label.config(text=f"Date: {current_date}")
        self.time_label.config(text=f"Time: {current_time}")

        if hasattr(self.root, 'message_label'):
            self.root.message_label.pack_forget()
        self.root.message_label = top_message_label

        # Display success message for a few seconds
        success_message = "Success!"
        self.display_success_message(success_message)

    def display_success_message(self, message):
        self.success_message_label.config(text=message, fg="blue")
        self.success_message_label.pack()
        self.root.after(3000, lambda: self.success_message_label.pack_forget())

    def stop_scanner(self):
        self.cap.release()
        self.root.destroy()

    def start_video_stream(self):
        def update_video_stream():
            ret, frame = self.cap.read()

            if ret:
                frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                photo = tk.PhotoImage(data=cv2.imencode('.ppm', frame)[1].tobytes())
                self.video_label.config(image=photo)
                self.video_label.photo = photo

                decoded_objects = decode(frame)
                for obj in decoded_objects:
                    qr_code_content = obj.data.decode('utf-8')
                    if qr_code_content not in self.recorded_qr_codes:
                        self.register_attendance(qr_code_content)

                self.video_label.after(10, update_video_stream)

        # Start the video stream in a separate thread
        video_thread = threading.Thread(target=update_video_stream)
        video_thread.daemon = True
        video_thread.start()

        self.root.mainloop()

if __name__ == "__main__":
    app = QRScannerApp()
