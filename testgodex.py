import serial 
import time
import os
COM_PORT = "/dev/ttyUSB0"
BAUD_RATE = 9600

label_data = """
^AT
^O0
^D0
^C1
^P1
^Q16.0,3.0
^W24
^L
RFW,H,2,24,1,00000001749121047830A347
W64,45,5,2,L,3,3,38,0
thuocsi.vn/qr/00000001749121047830A347
E
"""

def send_ezpl_file():
    try:
        with serial.Serial(COM_PORT, BAUD_RATE, timeout=1) as ser:
            ser.write(label_data.encode("ascii"))
            ser.flush()
            print(f"Sending... {COM_PORT}")
    except Exception as e:
            print(f"Error send file: {e}")

if __name__ == "__main__":
    send_ezpl_file()
