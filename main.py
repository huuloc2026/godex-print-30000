import pandas as pd
import os
import serial 
import time
import logging

# === Config ===
INPUT_FILE = "input/qr_60000.xlsx"
OUTPUT_DIR = "output"
ROWS_PER_FILE = 2000
IS_SIMULATION = False
LOG_FILE = "excel.log"
target = 60000  

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
thuocsi.vn/qr/{epc}
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
def send_ezpl_simulated(epc: str):
    label = label_template.format(epc=epc)
    logging.info(f"\n==== SIMULATED EZPL ====\n{label}\n=========================\n")

# === Real Print Mode ===
def send_ezpl_file(epc: str):
    label = label_template.format(epc=epc)

    try:
        with serial.Serial(COM_PORT, BAUD_RATE, timeout=1) as ser:
            ser.write(label.encode("ascii"))
            ser.flush()
        logging.info(f"✅ Sent label to printer: EPC={epc}")
        time.sleep(0.05) 
    except Exception as e:
        logging.error(f"❌ Error sending to printer: {e}")

# === Dispatcher ===
def send_ezpl(epc: str):
    if IS_SIMULATION:
        send_ezpl_simulated(epc)
    else:
        send_ezpl_file(epc)

def SplitExcel():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    logging.info(f"Loading Excel file: {INPUT_FILE}")
    df = pd.read_excel(INPUT_FILE, header=None)
    total_rows = len(df)
    if target != total_rows:
        logging.warning(f"⚠️ Expected {target} rows, but found {total_rows} rows in the file.")
        return
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
                epc = str(row[1])  # Cột A
               
                if pd.notna(epc):
                    send_ezpl(epc)
                    logging.info(f"{file_path} - idx - {idx} ✅ EPC {epc}")
                else:
                    logging.warning(f"⚠️ Missing EPC or QR at row {idx + 1}")
            except Exception as e:
                logging.error(f"❌ Error at row {idx + 1}: {e}")

    except Exception as e:
        logging.error(f"Failed to read Excel: {e}")



# === Example manual usage ===
if __name__ == "__main__":
    # SplitExcel()
    # ProcessLabelRange("output/split_part_01.xlsx", start=1, end=2)
    # ProcessLabelRange("output/split_part_01.xlsx", start=2, end=3)
    # ProcessLabelRange("output/split_part_01.xlsx", start=4, end=5)
    # ProcessLabelRange("output/split_part_01.xlsx", start=11, end=12)
    # ProcessLabelRange("output/split_part_01.xlsx", start=12, end=13)
    # ProcessLabelRange("output/split_part_01.xlsx", start=13, end=51)
    # ProcessLabelRange("output/split_part_01.xlsx", start=51, end=101)
    # ProcessLabelRange("output/split_part_01.xlsx", start=101, end=501)
    
    # ProcessLabelRange("output/split_part_01.xlsx", start=279, end=501)
    # ProcessLabelRange("output/split_part_01.xlsx", start=501, end=1001)
    # ProcessLabelRange("output/split_part_01.xlsx", start=1001, end=1051)
    # ProcessLabelRange("output/split_part_01.xlsx", start=1051, end=1071)
    # ProcessLabelRange("output/split_part_01.xlsx", start=1071, end=2001)
    # ProcessLabelRange("output/split_part_01.xlsx", start=1545, end=1546)
    # ProcessLabelRange("output/split_part_01.xlsx", start=1551, end=2001)
    
    # ProcessLabelRange("output/split_part_02.xlsx", start=1, end=11)
    # ProcessLabelRange("output/split_part_02.xlsx", start=11, end=50)
    # ProcessLabelRange("output/split_part_02.xlsx", start=50, end=1001)
    
    
    # ProcessLabelRange("output/split_part_02.xlsx", start=1001, end=2001)
    # KPI:   4000/50.000
    
    
    # ProcessLabelRange("output/split_part_03.xlsx", start=1, end=1001)
    # ProcessLabelRange("output/split_part_03.xlsx", start=1001, end=2001)
    
    
    
    # ProcessLabelRange("output/split_part_04.xlsx", start=1, end=501)
    # ProcessLabelRange("output/split_part_04.xlsx", start=501, end=601)
    # ProcessLabelRange("output/split_part_04.xlsx", start=601, end=602)
    
    # ProcessLabelRange("output/split_part_04.xlsx", start=602, end=1001)
    # ProcessLabelRange("output/split_part_04.xlsx", start=1001, end=2001)
    
    
    # ProcessLabelRange("output/split_part_05.xlsx", start=1, end=1001)
    # ProcessLabelRange("output/split_part_05.xlsx", start=1001, end=2001)
    
    # ProcessLabelRange("output/split_part_06.xlsx", start=1, end=1001)
    
    # ProcessLabelRange("output/split_part_06.xlsx", start=1001, end=2001) 
    
    # ProcessLabelRange("output/split_part_07.xlsx", start=1, end=2)

    # ProcessLabelRange("output/split_part_07.xlsx", start=2, end=1001)
    # ProcessLabelRange("output/split_part_07.xlsx", start=793, end=794)

    # ProcessLabelRange("output/split_part_07.xlsx", start=795, end=1001)
    
    # ProcessLabelRange("output/split_part_07.xlsx", start=1001, end=2001)
    
    # ProcessLabelRange("output/split_part_07.xlsx", start=1124, end=1125)
    # ProcessLabelRange("output/split_part_07.xlsx", start=1864, end=1865)
    
    # ProcessLabelRange("output/split_part_07.xlsx", start=1873, end=2001)

    # ProcessLabelRange("output/split_part_08.xlsx", start=1, end=1001)
    # ProcessLabelRange("output/split_part_08.xlsx", start=67, end=69)
    # ProcessLabelRange("output/split_part_08.xlsx", start=79, end=101)
    # ProcessLabelRange("output/split_part_08.xlsx", start=101, end=501)
    # ProcessLabelRange("output/split_part_08.xlsx", start=280, end=281)
    # ProcessLabelRange("output/split_part_08.xlsx", start=281, end=501)
    # ProcessLabelRange("output/split_part_08.xlsx", start=501, end=550)
    # ProcessLabelRange("output/split_part_08.xlsx", start=550, end=601)
    # ProcessLabelRange("output/split_part_08.xlsx", start=601, end=801)
    
    # ProcessLabelRange("output/split_part_08.xlsx", start=801, end=2001) 
    
    # ProcessLabelRange("output/split_part_09.xlsx", start=699, end=700)
    # ProcessLabelRange("output/split_part_09.xlsx", start=702, end=1001)
    # ProcessLabelRange("output/split_part_09.xlsx", start=1001, end=2001)
    # ProcessLabelRange("output/split_part_09.xlsx", start=1217, end=1220)
    # # ProcessLabelRange("output/split_part_09.xlsx", start=1268, end=1269)
    # ProcessLabelRange("output/split_part_09.xlsx", start=1272, end=1273)
    # ProcessLabelRange("output/split_part_09.xlsx", start=1297, end=1298)
    # ProcessLabelRange("output/split_part_09.xlsx", start=1297, end=2001)
    
    #01/07/2025
    # ProcessLabelRange("output/split_part_10.xlsx", start=1, end=2)  
    
    # ProcessLabelRange("output/split_part_10.xlsx", start=2, end=1001)  
#ProcessLabelRange("output/split_part_10.xlsx", start=1003, end=2001)


    # ProcessLabelRange("output/split_part_11.xlsx", start=1, end=1001)
    # ProcessLabelRange("output/split_part_11.xlsx", start=1788, end=2001)
    
    # ProcessLabelRange("output/split_part_12.xlsx", start=1, end=2)
        
    # ProcessLabelRange("output/split_part_12.xlsx", start=2, end=1001)
    # ProcessLabelRange("output/split_part_12.xlsx", start=444, end=1001)
    # ProcessLabelRange("output/split_part_12.xlsx", start=1001, end=2001)
    # ProcessLabelRange("output/split_part_13.xlsx", start=1, end=1001)
    
    # ProcessLabelRange("output/split_part_13.xlsx", start=602, end=603)
    # ProcessLabelRange("output/split_part_13.xlsx", start=603, end=1001)
    
    # ProcessLabelRange("output/split_part_13.xlsx", start=733, end=734)
    # ProcessLabelRange("output/split_part_13.xlsx", start=715, end=726)
    # ProcessLabelRange("output/split_part_13.xlsx", start=726, end=736)
    # ProcessLabelRange("output/split_part_13.xlsx", start=736, end=737)
    
    # ProcessLabelRange("output/split_part_13.xlsx", start=771, end=775)
    # ProcessLabelRange("output/split_part_13.xlsx", start=777, end=778)
    # ProcessLabelRange("output/split_part_13.xlsx", start=792, end=1001)
    # ProcessLabelRange("output/split_part_13.xlsx", start=1001, end=1006)
    # ProcessLabelRange("output/split_part_13.xlsx", start=1006, end=2001)
    
    # ProcessLabelRange("output/split_part_14.xlsx", start=161, end=1001)
    
    
    
    
    # ProcessLabelRange("output/split_part_14.xlsx", start=1001, end=2001)  
    
    # ProcessLabelRange("output/split_part_15.xlsx", start=1, end=1001)
    
    # ProcessLabelRange("output/split_part_15.xlsx", start=1001, end=2001)
    
    
    # ProcessLabelRange("output/split_part_16.xlsx", start=1, end=1001)
    
    # ProcessLabelRange("output/split_part_16.xlsx", start=1001, end=2001) 
    
       
    # ProcessLabelRange("output/split_part_17.xlsx", start=1, end=1001)
    # ProcessLabelRange("output/split_part_17.xlsx", start=1001, end=2001)
    
        
    # ProcessLabelRange("output/split_part_18.xlsx", start=1, end=1001)
    
    # ProcessLabelRange("output/split_part_18.xlsx", start=508, end=509)
    # ProcessLabelRange("output/split_part_18.xlsx", start=516, end=518)
    # ProcessLabelRange("output/split_part_18.xlsx", start=552,end=553)
    # ProcessLabelRange("output/split_part_18.xlsx", start=557, end=558)
    ProcessLabelRange("output/split_part_18.xlsx", start=559, end=1001)
    
    ################################################################
    ##                            NEXT                            ##
    ################################################################
    
    
   
    

    
    

    # ProcessLabelRange("output/split_part_19.xlsx", start=1, end=1001)
    # ProcessLabelRange("output/split_part_19.xlsx", start=1001, end=2001)
    
    # ProcessLabelRange("output/split_part_20.xlsx", start=1, end=1001)
    # ProcessLabelRange("output/split_part_20.xlsx", start=1001, end=2001)
    
    
    #ProcessLabel part 21 - 25
    
    # ProcessLabelRange("output/split_part_21.xlsx", start=1, end=1001)
    # ProcessLabelRange("output/split_part_21.xlsx", start=1001, end=2001)
    # ProcessLabelRange("output/split_part_22.xlsx", start=1, end=1001)
    # ProcessLabelRange("output/split_part_22.xlsx", start=1001, end=2001)
    # ProcessLabelRange("output/split_part_23.xlsx", start=1, end=1001)
    # ProcessLabelRange("output/split_part_23.xlsx", start=1001, end=2001)
    # ProcessLabelRange("output/split_part_24.xlsx", start=1, end=1001)
    # ProcessLabelRange("output/split_part_24.xlsx", start=1001, end=2001)
    # ProcessLabelRange("output/split_part_25.xlsx", start=1, end=1001)
    # ProcessLabelRange("output/split_part_25.xlsx", start=1001, end=2001)  
     
    
    
    
    
    
    
    
    
    
    
    
    



    