import pandas as pd
import re
from openpyxl import load_workbook

# 設定檔案名稱
input_file = 'Sunplus_Yield_control_table.xlsx'
output_file = 'yield_trend_4.xlsx'  # 更新檔案名稱
sheet_name = 'QAL642E LFBGA 487B'

# 指定要保留的欄位（使用你的格式）
columns_to_keep = "B, C, D, F, G, S, T"

try:
    # 讀取 Excel，篩選特定欄位，跳過第一列
    df = pd.read_excel(input_file, sheet_name=sheet_name, usecols=columns_to_keep, skiprows=1)

    # **1️⃣ 新增 RT rate 欄位**
    df["RT rate"] = None  # 預設值

    # **2️⃣ 解析 PGM Name，修改 FT 為 FT1、FT2...**
    def modify_ft(station, pgm_name):
        """ 如果 Station 是 FT，則從 PGM Name 中提取 f 後的數字，變成 FT1、FT2... """
        if station == "FT":
            match = re.search(r"f(\d+)", pgm_name)  # 尋找 'f' 後的數字
            if match:
                return f"FT{match.group(1)}"
        return station

    df["Station"] = df.apply(lambda row: modify_ft(row["Station"], row["PGM Name"]), axis=1)

    # **3️⃣ 計算 RT rate**
    rt_rate = None  # 變數用來存儲當前 RT rate

    for idx in df.index:
        station = str(df.at[idx, "Station"])

        if station.startswith("FT"):  # 每組的起點
            rt_rate = 0  # 初始化 RT rate
            rt_start_idx = idx  # 記錄該組起點索引

        elif re.match(r"R(\d+)", station):  # R1, R2, ..., RN
            rt_rate = max(rt_rate, int(re.match(r"R(\d+)", station).group(1)))

        elif station == "Total":  # 到達該組終點
            # 填入整組 FT ~ Total 的 RT rate
            df.loc[rt_start_idx:idx, "RT rate"] = rt_rate
            rt_rate = None  # 重置 RT rate

    # **4️⃣ 刪除包含 NaN 的列**
    df_cleaned = df.dropna()

    # **5️⃣ 儲存結果**
    df_cleaned.to_excel(output_file, index=False)

    # **6️⃣ 調整 Excel 欄位寬度**
    wb = load_workbook(output_file)
    ws = wb.active  # 取得預設的工作表

    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter  # 取得欄位的字母，例如 'A', 'B'
        
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass

        adjusted_width = max_length + 2  # 加 2 讓欄位更美觀
        ws.column_dimensions[col_letter].width = adjusted_width

    wb.save(output_file)  # 儲存 Excel 檔案

    print(f"✅ 資料已成功儲存到 {output_file}，欄位寬度已自動調整！")

except FileNotFoundError:
    print("❌ 找不到原始檔案，請檢查檔案名稱和路徑。")
except ValueError as e:
    print(f"❌ 發生錯誤: {e}")
except Exception as e:
    print(f"❌ 發生未知錯誤: {e}")
