import pandas as pd
import os
import serial 
import time
import logging

# === Config ===
INPUT_FILE = "input/qr_code_24.xlsx"
OUTPUT_DIR = "output"
ROWS_PER_FILE = 2000
IS_SIMULATION = True
LOG_FILE = "excel.log"
target = 50000  

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
        chunk_df.to_excel(output_path, index=False)
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
                epc = str(row[0])  # Cột A
                qr = str(row[3])   # Cột D
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
    # SplitExcel()
    ProcessLabelRange("output/split_part_01.xlsx", start=1, end=5)
    
    
    
    
    
    
    
    



    