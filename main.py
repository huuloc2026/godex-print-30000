import pandas as pd
import os
import serial 
import time
import logging

# === Config ===
INPUT_FILE = "input/qr_code_24.xlsx"
OUTPUT_DIR = "output"
ROWS_PER_FILE = 1000
IS_SIMULATION = False
LOG_FILE = "split_excel.log"


COM_PORT = "/dev/ttyUSB0"
BAUD_RATE = 9600
# === EZPL Template ===
label_template = """
^AT
^O0
^D0
^C1
^P1
^Q16.0,10.0
^W24
^L
RFW,H,2,24,1,{epc}
W64,45,5,2,L,3,3,38,0
{qr}
E
"""

# === Setup Logging ===
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler()
    ]
)
# === Simulation Mode ===
def send_ezpl_simulated(epc: str, qr: str):
    label = label_template.format(epc=epc, qr=qr)
    logging.info(f"\n==== SIMULATED EZPL ====\n{label}\n=========================\n")

# === Real Print Mode ===
def send_ezpl_file(epc: str, qr: str):
    label = label_template.format(epc=epc, qr=qr)

    try:
        with serial.Serial(COM_PORT, BAUD_RATE, timeout=1) as ser:
            ser.write(label.encode("ascii"))
            ser.flush()
        logging.info(f"✅ Sent label to printer: EPC={epc}, QR={qr}")
        # Small delay to avoid overloading printer buffer
        time.sleep(0.05) 
    except Exception as e:
        logging.error(f"❌ Error sending to printer: {e}")

# === Dispatcher ===
def send_ezpl(epc: str, qr: str):
    if IS_SIMULATION:
        send_ezpl_simulated(epc, qr)
    else:
        send_ezpl_file(epc, qr)

def SplitExcel():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    logging.info(f"Loading Excel file: {INPUT_FILE}")
    df = pd.read_excel(INPUT_FILE, header=None)
    total_rows = len(df)
    logging.info(f"Total rows loaded (excluding header): {total_rows}")

    num_chunks = (total_rows + ROWS_PER_FILE - 1) // ROWS_PER_FILE

    for i in range(num_chunks):
        start_row = i * ROWS_PER_FILE
        end_row = min((i + 1) * ROWS_PER_FILE, total_rows)
        chunk_df = df.iloc[start_row:end_row]

        output_path = os.path.join(OUTPUT_DIR, f"split_part_{i+1:02d}.xlsx")
        chunk_df.to_excel(output_path, index=True)
        logging.info(f"Saved rows {start_row + 1}–{end_row} to {output_path}")

    logging.info(f"✅ Done. {num_chunks} files created in {OUTPUT_DIR}")

# def ProcessLabelsFromFile(file_path: str, batch_size: int = 200):
#     if not os.path.exists(file_path):
#         logging.error(f"Missing file: {file_path}")
#         return

#     logging.info(f"Processing file: {file_path}")
#     try:
#         df = pd.read_excel(file_path, header=None)

#         total_rows = len(df)
#         total_batches = (total_rows + batch_size - 1) // batch_size

#         for batch_num in range(total_batches):
#             start_idx = batch_num * batch_size
#             end_idx = min(start_idx + batch_size, total_rows)

#             logging.info(f"📦 Printing batch {batch_num + 1}/{total_batches} (rows {start_idx + 1} to {end_idx})")

#             batch_df = df.iloc[start_idx:end_idx]
#             for idx, row in batch_df.iterrows():
#                 try:
#                     epc = str(row[1])  # Column B
#                     qr = str(row[4])   # Column E
#                     if pd.notna(epc) and pd.notna(qr):
#                         send_ezpl(epc, qr)
#                         logging.info(f"✅ Success EPC: {epc}")
#                     else:
#                         logging.warning(f"⚠️ Skipping row {idx + 1}: missing EPC or QR")
#                 except Exception as e:
#                     logging.error(f"❌ Error at row {idx + 1}: {e}")

#     except Exception as e:
#         logging.error(f"Failed to read {file_path}: {e}")

def ProcessLabelRange(file_path: str, start: int, end: int):
    """
    In nhãn từ dòng `start` đến `end` (1-based index, giống Excel).
    """
    if not os.path.exists(file_path):
        logging.error(f"Missing file: {file_path}")
        return

    logging.info(f"📄 Processing rows {start} to {end} from: {file_path}")

    try:
        df = pd.read_excel(file_path, header=None)
        total_rows = len(df)

        # Chuyển start/end từ 1-based (Excel style) sang 0-based (pandas)
        start_idx = start
        end_idx = min(end, total_rows)

        for idx, row in df.iloc[start_idx:end_idx].iterrows():
            try:
                epc = str(row[1])  # Cột B
                qr = str(row[4])   # Cột E
                if pd.notna(epc) and pd.notna(qr):
                    send_ezpl(epc, qr)
                    logging.info(f"idx - {idx} ✅ EPC {epc}")
                else:
                    logging.warning(f"⚠️ Missing EPC or QR at row {idx + 1}")
            except Exception as e:
                logging.error(f"❌ Error at row {idx + 1}: {e}")

    except Exception as e:
        logging.error(f"Failed to read Excel: {e}")



# === Example manual usage ===
if __name__ == "__main__":
    # Uncomment below to split
    # SplitExcel()

    # Manually print one file (e.g., split_part_01.xlsx)
    # test_file = os.path.join(OUTPUT_DIR, "split_part_01.xlsx")
    # ProcessLabelRange("output/split_part_01.xlsx", start=1, end=10)

    # ProcessLabelRange("output/split_part_01.xlsx", start=1, end=10)
    # ProcessLabelRange("output/split_part_01.xlsx", start=10, end=20)
    # ProcessLabelRange("output/split_part_01.xlsx", start=20, end=30)
    # ProcessLabelRange("output/split_part_01.xlsx", start=30, end=80)
    # ProcessLabelRange("output/split_part_01.xlsx", start=80, end=101)
    # ProcessLabelRange("output/split_part_01.xlsx", start=101, end=301)
    # ProcessLabelRange("output/split_part_01.xlsx", start=301, end=601)
    ###ProcessLabelRange("output/split_part_01.xlsx", start=601, end=901)

    # ProcessLabelRange("output/split_part_01.xlsx", start=753, end=1001)
    

    #FILE 2 - DONE  
    #ProcessLabelRange("output/split_part_02.xlsx", start=1, end=501)
    # ProcessLabelRange("output/split_part_02.xlsx", start=501, end=901)
    

    #FILE 3 - DONE 
    # ProcessLabelRange("output/split_part_03.xlsx", start=1, end=501)
    # ProcessLabelRange("output/split_part_03.xlsx", start=501, end=1001)

    #FILE 4 - DONE - 4000
    # ProcessLabelRange("output/split_part_04.xlsx", start=1, end=501)
    # ProcessLabelRange("output/split_part_04.xlsx", start=501, end=1001)


    #10/06/2025
    #FILE 5 - DONE - 5000 - ( KPI: 1000 / 12000 )
    # ProcessLabelRange("output/split_part_05.xlsx", start=1, end=2)
    # ProcessLabelRange("output/split_part_05.xlsx", start=2, end=501)
    # ProcessLabelRange("output/split_part_05.xlsx", start=501, end=1001)

    #FILE 6 - 6000 - ( KPI: 2000 / 12000 )
    # ProcessLabelRange("output/split_part_06.xlsx", start=1, end=2)
    # ProcessLabelRange("output/split_part_06.xlsx", start=2, end=501) # 5500 tag
    # ProcessLabelRange("output/split_part_06.xlsx", start=501, end=1001)

    #FILE 7 - 7000 - ( KPI: 3000 / 12000 ) 12h
    # ProcessLabelRange("output/split_part_07.xlsx", start=1, end=2)
    # ProcessLabelRange("output/split_part_07.xlsx", start=2, end=3)
    # ProcessLabelRange("output/split_part_07.xlsx", start=3, end=501)
    # ProcessLabelRange("output/split_part_07.xlsx", start=501, end=1001)



    #FILE 8 - 8000 - ( KPI: 4000 / 12000 )


    # ProcessLabelRange("output/split_part_08.xlsx", start=1, end=501) 7500/20000
    # ProcessLabelRange("output/split_part_08.xlsx", start=501, end=1001) #8000/20000

    #FILE 9 - 9000 - ( KPI: 5000 / 12000 )

    # ProcessLabelRange("output/split_part_09.xlsx", start=1, end=2) 

    # ProcessLabelRange("output/split_part_09.xlsx", start=2, end=1001) #9000/20000


    #FILE 10 - 9000 - ( KPI: 6000 / 12000 )
    # ProcessLabelRange("output/split_part_10.xlsx", start=1, end=1001) #10000/20000

    #FILE 11 - 10000 - ( KPI: 6000 / 12000 )
    # ProcessLabelRange("output/split_part_11.xlsx", start=1, end=1001) #11000/20000
    
    
    #17/06/2025
    # ProcessLabelRange("output/split_part_21.xlsx", start=1, end=11) 
    # ProcessLabelRange("output/split_part_21.xlsx", start=11, end=501) 
    # ProcessLabelRange("output/split_part_21.xlsx", start=501, end=1001) 
    # ProcessLabelRange("output/split_part_22.xlsx", start=1, end=501) 
    
    
    #18/06/2025 - 2000/10000
    # ProcessLabelRange("output/split_part_22.xlsx", start=501, end=502) 
    # ProcessLabelRange("output/split_part_22.xlsx", start=502, end=1001) # het muc
    # ProcessLabelRange("output/split_part_22.xlsx", start=580, end=581)
    # ProcessLabelRange("output/split_part_22.xlsx", start=581, end=582)
    # ProcessLabelRange("output/split_part_22.xlsx", start=582, end=1001)
    
    #3000/10000
    # ProcessLabelRange("output/split_part_23.xlsx", start=1, end=2)
    # ProcessLabelRange("output/split_part_23.xlsx", start=2, end=501)
    # ProcessLabelRange("output/split_part_23.xlsx", start=501, end=1001)
    
    #4000/10000
    # ProcessLabelRange("output/split_part_24.xlsx", start=1, end=2)
    # ProcessLabelRange("output/split_part_24.xlsx", start=1, end=501)
    # ProcessLabelRange("output/split_part_24.xlsx", start=501, end=601)
    # ProcessLabelRange("output/split_part_24.xlsx", start=601, end=1001)
    # ProcessLabelRange("output/split_part_24.xlsx", start=795, end=1001)
    
    #5000/10000
    # ProcessLabelRange("output/split_part_25.xlsx", start=1, end=501)
    # ProcessLabelRange("output/split_part_25.xlsx", start=501, end=1001)
     
    #6000/10000
    # ProcessLabelRange("output/split_part_26.xlsx", start=1, end=501)
    # ProcessLabelRange("output/split_part_26.xlsx", start=501, end=1001)
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    #8000/10000
    # ProcessLabelRange("output/split_part_27.xlsx", start=1, end=501)
    # ProcessLabelRange("output/split_part_27.xlsx", start=501, end=1001)
     
    # ProcessLabelRange("output/split_part_28.xlsx", start=163, end=1001)
    #==========
    
    
    #10000/10000
    # ProcessLabelRange("output/split_part_29.xlsx", start=3, end=1001)
 
    # ProcessLabelRange("output/split_part_30.xlsx", start=1, end=1001)
    # ProcessLabelRange("output/split_part_30.xlsx", start=150, end=151)
    ProcessLabelRange("output/split_part_30.xlsx", start=151, end=1001)
    
    
    
    
    
    
    
    



    