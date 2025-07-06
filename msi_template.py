import os
import sys
from dotenv import load_dotenv
import pandas as pd
from scipy.stats import norm

def log(msg):
    print(f"[msi_template] {msg}", file=sys.stderr)

def compute_metrics(series):
    # a) Freq kategori 1–5
    freq = series.value_counts().reindex(range(1,6), fill_value=0)
    N    = freq.sum()
    # b) Proporsi
    P    = freq / N
    # c) Cumulative P
    CP   = P.cumsum()
    # d) MID_CP
    MID  = CP - P/2
    # e) 0.5 – MID_CP
    hm   = (0.5 - MID)
    # f) Z (clipped ±3.9)
    Z    = MID.apply(lambda x: -3.9 if x==0 else (3.9 if x==1 else norm.ppf(x))).tolist()
    # g) ZC = Z – min(Z)
    min_z = min(Z)
    ZC    = [z - min_z for z in Z]
    # h) Pembulatan ZC
    ZCr   = [int(round(v)) for v in ZC]
    return (
        freq.tolist(),
        P.tolist(),
        CP.tolist(),
        MID.tolist(),
        hm.tolist(),
        Z,
        ZC,
        ZCr
    )

def main():
    # 1. Load .env
    load_dotenv()
    INPUT_FILE    = os.getenv("INPUT_FILE",   "responses.xlsx")
    INPUT_SHEET   = os.getenv("INPUT_SHEET",  "Sheet1")
    OUTPUT_FILE   = os.getenv("OUTPUT_FILE",  "MSI_all_in_one.xlsx")
    REVERSE_ENV   = os.getenv("REVERSE_ITEMS", "")
    REVERSE_ITEMS = [i.strip() for i in REVERSE_ENV.split(",") if i.strip()]

    log(f"Config → INPUT_FILE='{INPUT_FILE}', INPUT_SHEET='{INPUT_SHEET}', OUTPUT_FILE='{OUTPUT_FILE}'")
    log(f"Reverse-coded items: {REVERSE_ITEMS or '(none)'}")

    # 2. Read data
    log(f"Membaca data dari '{INPUT_FILE}' sheet '{INPUT_SHEET}' …")
    df = pd.read_excel(INPUT_FILE, sheet_name=INPUT_SHEET)
    log(f"Data shape: {df.shape[0]} rows × {df.shape[1]} columns")
    items = list(df.columns)
    log(f"Detected items ({len(items)}): {items[:5]} … {items[-5:]}")

    # 3. Build Sheet1 rows
    log("Menghitung metrik per item untuk Sheet1 …")
    rows = []
    for idx, item in enumerate(items, 1):
        s = df[item]
        if item in REVERSE_ITEMS:
            log(f"  • Reverse-coding '{item}'")
            s = s.apply(lambda x: 6 - x)

        freq, P, CP, MID, hm, Z, ZC, ZCr = compute_metrics(s)
        rows.append([item, 1,2,3,4,5])
        rows.append(['F'         , *freq])
        rows.append(['P=F/N'     , *P])
        rows.append(['CP'        , *CP])
        rows.append(['MID_CP'    , *MID])
        rows.append(['0.5-MID_CP', *hm])
        rows.append(['Z'         , *Z])
        rows.append(['ZC'        , *ZC])
        rows.append(['Pembulatan', *ZCr])
        rows.append([''] * 6)  # blank
        rows.append([''] * 6)  # blank
        if idx % 10 == 0 or idx == len(items):
            log(f"  • Processed {idx}/{len(items)} items")

    sheet1_df = pd.DataFrame(rows)
    log(f"Sheet1 built: {sheet1_df.shape[0]} rows")

    # 4. Build Sheet2 in one go to avoid fragmentation
    log("Membuat Sheet2 (Pembulatan per responden) …")
    data2 = {}
    for item in items:
        s = df[item]
        if item in REVERSE_ITEMS:
            s = s.apply(lambda x: 6 - x)
        _,_,_,_,_,_,_, ZCr = compute_metrics(s)
        mapping = {i+1: ZCr[i] for i in range(5)}
        data2[item] = s.map(mapping)
    sheet2_df = pd.DataFrame(data2)
    log(f"Sheet2 built: {sheet2_df.shape[0]} rows × {sheet2_df.shape[1]} columns")

    # 5. Write to Excel
    log(f"Menulis output ke '{OUTPUT_FILE}' …")
    with pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter') as writer:
        sheet1_df.to_excel(writer, sheet_name='Sheet1', index=False, header=False)
        sheet2_df.to_excel(writer, sheet_name='Sheet2', index=False)
    log("Selesai menulis file.")

if __name__ == "__main__":
    main()