import xlwings as xw
import os
from glob import glob

# 合併後的輸出檔案
output_file = "merged_yield_trend.xlsx"
if os.path.exists(output_file):
    os.remove(output_file)

# 啟動 Excel 應用程式
app = xw.App(visible=False)
merged_wb = app.books.add()

# 尋找所有待合併的 Excel 檔案
files = [f for f in glob("*_yield_trend.xlsx") if os.path.basename(f) != output_file]

for file in files:
    prefix = os.path.basename(file).replace("_yield_trend.xlsx", "")
    print(f"📥 處理檔案：{file}")

    src_wb = app.books.open(file)

    for sheet in src_wb.sheets:
        # 複製工作表到 merged_wb
        sheet.api.Copy(Before=merged_wb.sheets[0].api)

        # 重新命名複製的工作表（在最前面）
        copied_sheet = merged_wb.sheets[0]
        new_name = f"{prefix}_{sheet.name}"[:31]  # 限制在 Excel 的工作表名上限
        copied_sheet.name = new_name

        print(f"  ➜ 加入工作表：{new_name}")

    src_wb.close()

# 刪除預設空白工作表
if len(merged_wb.sheets) > 1:
    try:
        merged_wb.sheets[-1].delete()
    except:
        pass

# 儲存並關閉
merged_wb.save(output_file)
merged_wb.close()
app.quit()

print(f"\n✅ 合併完成，儲存為：{output_file}")
