import pandas as pd
from openpyxl import Workbook, load_workbook

"""
wf_cl_df = "D:/00_herne/_export_SmartSheet/WF_CONTROL_LIST_20200923.xlsx"   # work file control list data frame

data_file = "D:/00_herne/_export_SmartSheet/IZ224_Weld_Map.xlsx"
df = pd.read_excel(data_file, sheet_name="IZ224_Weld_Map")

xl_file = "D:/00_herne/00_template/WF_template.xlsx"
wb = load_workbook(xl_file, read_only=False)
writer = pd.ExcelWriter(xl_file, engine="openpyxl")
writer.book = wb
writer.sheets = dict((ws.title, ws) for ws in wb.worksheets)

df["counts"] = df.groupby(["Line KKS"])["W_NO"].transform("count")
df_filtered = pd.DataFrame(df.filter(["Line KKS", "counts"], axis=1).drop_duplicates())
"""
df = pd.read_json('data_file.json')


# print(df[df["KKS"] == "60EGC10BR001"])
# print(df["KKS_connected"])
# print(df_filtered)

# df[df["Line KKS"] == "60LBA20BR004"].to_excel(writer, sheet_name="04_Matrix", startrow=5, header=False, index=False, float_format="%.2f")
# writer.save()

for i in df["KKS_connected"]:
    print(i)