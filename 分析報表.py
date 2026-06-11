import os
import pandas as pd
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import TkinterDnD, DND_FILES
import threading

data_path = None

def select_file():
    filename = filedialog.askopenfilename(
        title="請選擇Excel檔",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if filename:
        file_var.set(filename)

def drop(event):
    file_var.set(event.data.strip("{}"))

def start_process():
    if not file_var.get().strip():
        messagebox.showerror("錯誤", "請選擇Excel檔")
        return

    progress["value"] = 0
    status_label.config(text="準備開始...")

    threading.Thread(target=run_analysis, daemon=True).start()


root = TkinterDnD.Tk()
root.title("分公司線路使用率分析")
root.geometry("820x260")

tk.Label(root, text="拖曳 Excel 或按瀏覽選檔", font=("Microsoft JhengHei", 12)).pack(pady=8)

file_var = tk.StringVar()

entry = tk.Entry(root, textvariable=file_var, width=110)
entry.pack(padx=10)

entry.drop_target_register(DND_FILES)
entry.dnd_bind("<<Drop>>", drop)

tk.Button(root, text="瀏覽檔案", command=select_file).pack(pady=5)

tk.Button(
    root,
    text="開始分析",
    bg="#28a745",
    fg="white",
    width=18,
    command=start_process
).pack(pady=5)

progress = ttk.Progressbar(root, orient="horizontal", length=700, mode="determinate")
progress.pack(pady=10)

status_label = tk.Label(root, text="等待開始")
status_label.pack()

def process_sheet(sheet_name):
    df = pd.read_excel(data_path, sheet_name=sheet_name, header=None)
    df = df.ffill(axis=1)

    result = []

    def find_row(keyword):
        for i, val in enumerate(df.iloc[:, 0]):
            if pd.notna(val) and keyword in str(val):
                return i
        return None

    avg_row = find_row("平均")
    max_row = find_row("最大值")
    bw_row = find_row("頻寬")

    if max_row is None or bw_row is None:
        raise ValueError(f"{sheet_name} 找不到 最大值 或 頻寬")

    data_end_row = avg_row if avg_row is not None else max_row
    data_start_row = 3

    for col in range(1, df.shape[1]):

        device = str(df.iloc[1, col]).strip()
        interface = str(df.iloc[2, col]).strip()

        bandwidth = pd.to_numeric(df.iloc[bw_row, col], errors='coerce')
        max_value = pd.to_numeric(df.iloc[max_row, col], errors='coerce')

        if pd.isna(bandwidth) or pd.isna(max_value):
            continue

        exceed_dates = []
        warn90_dates = []

        for row in range(data_start_row, data_end_row):
            value = pd.to_numeric(df.iloc[row, col], errors='coerce')
            date = df.iloc[row, 0]

            if pd.notna(value):
                if value > bandwidth:
                    exceed_dates.append(str(date))
                elif value > bandwidth * 0.9:
                    warn90_dates.append(str(date))

        if exceed_dates or warn90_dates:
            result.append({
                "設備名稱": device,
                "介面": interface,
                "最大值": max_value,
                "頻寬": bandwidth,
                "滿載天數": len(exceed_dates),
                "滿載日期": ", ".join(exceed_dates),
                "90%天數": len(warn90_dates),
                "90%日期": ", ".join(warn90_dates)
            })

    return pd.DataFrame(result, columns=[
        "設備名稱","介面","最大值","頻寬",
        "滿載天數","滿載日期","90%天數","90%日期"
    ])


def filter_oa(df):
    if df.empty:
        return df
    return df[
        (df["設備名稱"].str.contains("R", na=False)) &
        (df["設備名稱"].str.contains("B", na=False)) &
        (df["介面"].str.contains("0/0/1", na=False))
    ]


def filter_letter(df, letter):
    if df.empty:
        return df
    return df[df["設備名稱"].str.contains(letter, na=False)]


def run_analysis():
    global data_path

    try:
        data_path = file_var.get().strip()

        if not os.path.exists(data_path):
            messagebox.showerror("錯誤", "檔案不存在")
            return

        status_label.config(text="分析 Rx...")
        progress["value"] = 10

        rx_df = process_sheet("每日Rx最大值(Mbps)")

        status_label.config(text="分析 Tx...")
        progress["value"] = 40

        tx_df = process_sheet("每日Tx最大值(Mbps)")

        status_label.config(text="分類資料...")
        progress["value"] = 60

        oa_rx = filter_oa(rx_df)
        oa_tx = filter_oa(tx_df)

        rx_remaining = rx_df[~rx_df.index.isin(oa_rx.index)]
        tx_remaining = tx_df[~tx_df.index.isin(oa_tx.index)]

        a_rx = filter_letter(rx_remaining, "A")
        a_tx = filter_letter(tx_remaining, "A")
        b_rx = filter_letter(rx_remaining, "B")
        b_tx = filter_letter(tx_remaining, "B")

        status_label.config(text="輸出 Excel...")
        progress["value"] = 80

        output_path = os.path.join(
            os.path.dirname(data_path),
            f"分公司線路使用_in90%_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx"
        )

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            if not a_tx.empty:
                a_tx.to_excel(writer, sheet_name="A-Tx", index=False)
            if not a_rx.empty:
                a_rx.to_excel(writer, sheet_name="A-Rx", index=False)
            if not oa_tx.empty:
                oa_tx.to_excel(writer, sheet_name="OA-Tx", index=False)
            if not oa_rx.empty:
                oa_rx.to_excel(writer, sheet_name="OA-Rx", index=False)
            if not b_tx.empty:
                b_tx.to_excel(writer, sheet_name="OA2-Tx", index=False)
            if not b_rx.empty:
                b_rx.to_excel(writer, sheet_name="OA2-Rx", index=False)

        progress["value"] = 100
        status_label.config(text="完成")

        messagebox.showinfo("完成", f"輸出完成：\n{output_path}")

        try:
            os.startfile(output_path)
        except:
            pass

    except Exception as e:
        messagebox.showerror("錯誤", str(e))

root.mainloop()
