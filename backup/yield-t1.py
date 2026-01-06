import pandas as pd

# 設定檔案名稱
input_file = 'Sunplus_Yield_control_table.xlsx'
output_file = 'yield_trend.xlsx'
sheet_name = 'QAL642E LFBGA 487B'

# 指定要保留的欄位
columns_to_keep = "B, C, D, F, G, S, T"

try:
    # 讀取 Excel 並過濾欄位，跳過第一列
    df = pd.read_excel(input_file, sheet_name=sheet_name, usecols=columns_to_keep, skiprows=1)
    print("讀取並跳過第一列後的資料：")
    print(df.head(10))  # 先印出前 10 筆確認
    df.to_excel('yield_trend_0.xlsx')

  
   
except FileNotFoundError:
    print("❌ 找不到原始檔案，請檢查檔案名稱和路徑。")
except ValueError as e:
    print(f"❌ 發生錯誤: {e}")
except Exception as e:
    print(f"❌ 發生未知錯誤: {e}")
