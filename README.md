# QR Attendance Management System

This project consists of two Python scripts that create a complete system for managing QR-based attendance tracking, including QR code generation and scanning capabilities.

## Script 1: QR Code Email Sender with GUI

This script facilitates attendance management by generating unique QR codes for each student and emailing them with event details.

### Features:
1. **User Input Interface**: 
   - Uses Tkinter to create a GUI for collecting the sender’s email, password, event time, and venue details.
   - Allows loading of an Excel file containing student data.

2. **QR Code Generation**: 
   - Generates a QR code for each student containing their name and ID.

3. **Email Sending**: 
   - Sends each QR code to the respective student with event details and instructions.

4. **Excel Data Processing**: 
   - Reads student details from an Excel file for bulk QR code generation and emailing.

5. **Status Tracking**: 
   - Displays a list of successfully sent emails and counts within the Tkinter interface.

## Script 2: QR Code Scanner and Attendance Logger

This script sets up a QR code scanner application for recording attendance in real-time.

### Features:
1. **Camera Setup**: 
   - Uses OpenCV to access the camera feed and continuously scans for QR codes.

2. **QR Code Decoding and Verification**: 
   - Decodes QR codes using the Pyzbar library and verifies content for valid attendance information.

3. **Attendance Logging**: 
   - Records student details, current time, and date in an Excel sheet upon scanning valid QR codes.

4. **Real-Time Feedback**: 
   - Displays messages confirming successful attendance recordings with date and time.

5. **GUI Controls**: 
   - Offers options to stop the scanner or open the folder where attendance records are saved.

## Usage
1. Ensure you have the required libraries installed. You can install them using:
   ```bash
   pip install qrcode pandas openpyxl pyzbar opencv-python


QR Emailer image

![QR Sender Email image](https://github.com/user-attachments/assets/f2c53f48-8045-4774-90ac-8888b9193a77)

QR Code Attendance Scanner

![image](https://github.com/user-attachments/assets/0f50c1da-52b5-4482-95c3-da445bc55a91)

