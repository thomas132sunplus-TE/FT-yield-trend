import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.chart.axis import ChartLines
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.colors import ColorChoice
from openpyxl.chart.shapes import GraphicalProperties

# 設定檔案名稱
input_file = 'Sunplus_Yield_control_table.xlsx'
output_file = 'yield_trend_6.xlsx'  # 更新最終 Excel 檔案名稱
sheet_name = 'QAL642E LFBGA 487B'

# 指定要保留的欄位
columns_to_keep = "B, C, D, F, G, S, T"

try:
    # **1️⃣ 讀取 Excel，篩選特定欄位，跳過第一列**
    df = pd.read_excel(input_file, sheet_name=sheet_name, usecols=columns_to_keep, skiprows=1)

    # **2️⃣ 新增 RT rate 欄位**
    df["RT rate"] = None  # 預設值

    # **3️⃣ 解析 PGM Name，修改 FT 為 FT1、FT2...**
    def modify_ft(station, pgm_name):
        """ 如果 Station 是 FT，則從 PGM Name 中提取 f 後的數字，變成 FT1、FT2... """
        if station == "FT":
            match = re.search(r"f(\d+)", pgm_name)  # 尋找 'f' 後的數字
            if match:
                return f"FT{match.group(1)}"
        return station

    df["Station"] = df.apply(lambda row: modify_ft(row["Station"], row["PGM Name"]), axis=1)

    # **4️⃣ 計算 RT rate**
    rt_rate = None  # 變數用來存儲當前 RT rate

    for idx in df.index:
        station = str(df.at[idx, "Station"])

        if station.startswith("FT"):  # 每組的起點
            rt_rate = 0  # 初始化 RT rate
            rt_start_idx = idx  # 記錄該組起點索引

        elif re.match(r"R(\d+)", station):  # R1, R2, ..., RN
            rt_rate = max(rt_rate, int(re.match(r"R(\d+)", station).group(1)))

        elif station == "Total":  # 到達該組終點
            df.loc[rt_start_idx:idx, "RT rate"] = rt_rate
            rt_rate = None  # 重置 RT rate

    # **5️⃣ 刪除包含 NaN 的列**
    df_cleaned = df.dropna()

    # **6️⃣ 分類 `FT1`, `FT2`, `FT3` 到不同的 Sheet**
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for ft_group in df_cleaned["Station"].unique():
            if ft_group.startswith("FT"):
                df_cleaned[df_cleaned["Station"] == ft_group].to_excel(writer, sheet_name=ft_group, index=False)

    # **7️⃣ 調整 Excel 欄位寬度**
    wb = load_workbook(output_file)
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for col in ws.columns:
            max_length = max((len(str(cell.value)) for cell in col if cell.value), default=10)
            ws.column_dimensions[col[0].column_letter].width = max_length + 2

    # **8️⃣ 統一 `RT rate` Y 軸高度**
    max_rt_rate = max(df_cleaned["RT rate"])

    # **9️⃣ 為每個 FT Sheet 加入趨勢圖**
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        headers = [cell.value for cell in ws[1]]
        lot_col = headers.index("Lot#") + 1
        first_pass_col = headers.index("First Pass Yield") + 1
        overall_col = headers.index("Overall Yield") + 1
        rt_rate_col = headers.index("RT rate") + 1
        last_row = ws.max_row

        # **建立折線圖**
        combo_chart = LineChart()
        combo_chart.title = f"{sheet_name} - Yield & RT rate 趨勢"
        combo_chart.x_axis.title = "Lot#"
        combo_chart.y_axis.title = "Yield (%)"

        x_values = Reference(ws, min_col=lot_col, min_row=2, max_row=last_row)

        for col_index, series_name in [(first_pass_col, "First Pass Yield"), (overall_col, "Overall Yield")]:
            y_values = Reference(ws, min_col=col_index, min_row=1, max_row=last_row)
            combo_chart.add_data(y_values, titles_from_data=True)

        combo_chart.set_categories(x_values)

        # **建立柱狀圖**
        bar_chart = BarChart()
        bar_chart.y_axis.title = "RT rate"
        bar_chart.y_axis.axId = 200
        bar_chart.y_axis.majorGridlines = None  # 移除格線

        y_values = Reference(ws, min_col=rt_rate_col, min_row=1, max_row=last_row)
        bar_chart.add_data(y_values, titles_from_data=True)
        bar_chart.set_categories(x_values)

        combo_chart.y_axis.crosses = "max"
        combo_chart += bar_chart

        # **統一 `RT rate` 高度**
        bar_chart.y_axis.scaling.min = 0
        bar_chart.y_axis.scaling.max = max_rt_rate * 1.2

        # **淡化主要格線**
        gray_gridlines = ChartLines()
        gray_gridlines.spPr = GraphicalProperties()
        gray_gridlines.spPr.ln = LineProperties(solidFill=ColorChoice(prstClr="ltGray"))
        combo_chart.y_axis.majorGridlines = gray_gridlines

        # **放大圖表**
        combo_chart.width = 24
        combo_chart.height = 12

        # **📌 將圖表插入位置往左移一欄**
        ws.add_chart(combo_chart, "K5")  # 原本是 "L5"，現在改成 "K5"

    wb.save(output_file)
    print(f"✅ 資料已成功儲存到 {output_file}，圖表位置已左移一欄！")

except FileNotFoundError:
    print("❌ 找不到原始檔案，請檢查檔案名稱和路徑。")
except ValueError as e:
    print(f"❌ 發生錯誤: {e}")
except Exception as e:
    print(f"❌ 發生未知錯誤: {e}")
