import pandas as pd
import re

# 設定檔案名稱
input_file = 'Sunplus_Yield_control_table.xlsx'
output_file = 'yield_trend_1.xlsx'
sheet_name = 'QAL642E LFBGA 487B'

# 指定要保留的欄位（使用你的正確格式）
columns_to_keep = "B, C, D, F, G, S, T"

try:
    # 讀取 Excel，篩選特定欄位，跳過第一列
    df = pd.read_excel(input_file, sheet_name=sheet_name, usecols=columns_to_keep, skiprows=1)

    # 列印 DataFrame 欄位名稱，確保正確讀取
    print("📌 DataFrame 欄位名稱:", df.columns.tolist())

    # **1️⃣ 新增 RT rate 欄位**
    df["RT rate"] = None  # 預設值

    # **2️⃣ 分組處理 FT 到 Total 之間的 Station**
    rt_rate = None  # 變數用來存儲當前 RT rate

    for idx in df.index:
        station = str(df.at[idx, "Station"])

        if station == "FT":  # 每組的起點
            rt_rate = 0  # 初始化 RT rate
            rt_start_idx = idx  # 記錄該組起點索引

        elif re.match(r"R(\d+)", station):  # R1, R2, ..., RN
            rt_rate = max(rt_rate, int(re.match(r"R(\d+)", station).group(1)))

        elif station == "Total":  # 到達該組終點
            # 填入整組 FT ~ Total 的 RT rate
            df.loc[rt_start_idx:idx, "RT rate"] = rt_rate
            rt_rate = None  # 重置 RT rate

    # **3️⃣ 儲存結果**
    df.to_excel(output_file, index=False)
    print(f"✅ 資料已成功儲存到 {output_file}")

except FileNotFoundError:
    print("❌ 找不到原始檔案，請檢查檔案名稱和路徑。")
except ValueError as e:
    print(f"❌ 發生錯誤: {e}")
except Exception as e:
    print(f"❌ 發生未知錯誤: {e}")
