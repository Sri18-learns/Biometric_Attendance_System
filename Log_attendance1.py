from datetime import datetime
import time
from openpyxl import Workbook, load_workbook
import serial

# Serial port setup
ser = serial.Serial('COM4', 9600)
#print("Serial port initialized:", ser.name)

# File configuration
file_name = "Log_attendance1.xlsx"

# Initialize attendance log and file
attendance_log = {}

try:
    # Load existing workbook or create a new one
    workbook = load_workbook(file_name)
    sheet = workbook.active
    print(f"Loaded existing file: {file_name}")
except FileNotFoundError:
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["ID", "In Time", "Out Time"])  # Add headers
    workbook.save(file_name)
    print(f"New file created: {file_name}")

print("Listening for serial data...")

try:
    ctrlcounter = 1
    #while True:
    while ctrlcounter == 1:
        if ser.in_waiting > 0:
            # Read and decode serial data
            line = ser.readline().decode('utf-8', errors='ignore').strip()
            #print("Received:", line)

            # Process valid attendance data
            if "Found ID #" in line:
                # Extract ID from the message
                id = line.split("#")[1].split()[0].strip()
                #timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                timestamp = datetime.now().strftime('%Y-%m-%d %H:%M')
                # Log attendance
                if id not in attendance_log or attendance_log[id]["state"] == "out":
                    # Log Time In
                    attendance_log[id] = {"state": "in", "last_time": timestamp}
                    sheet.append([id, timestamp, ""])  # Time In
                    print(f"ID {id}: Time logged at {timestamp}")
                    time.sleep(2)
                else:
                    # Log Time Out
                    last_time_in = attendance_log[id]["last_time"]
                    attendance_log[id] = {"state": "out", "last_time": timestamp}
                    sheet.append([id, last_time_in, timestamp])  # Time Out
                    print(f"ID {id}: Time Out logged at {timestamp}")
                    time.sleep(2)

                # Save changes to the Excel file
                workbook.save(file_name)
                print("Workbook saved successfully.")

except KeyboardInterrupt:
    print("Program terminated.")
except Exception as e:
    print(f"An error occurred: {e}")
finally:
    workbook.save(file_name)
    print(f"Final file save completed: {file_name}")


    
