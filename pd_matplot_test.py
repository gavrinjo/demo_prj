import pandas as pd
import matplotlib.pyplot as plt
from matplotlib import cycler

IPython_default = plt.rcParams.copy()

colors = cycler('color',
                ['#EE6666', '#3388BB', '#9988DD',
                 '#EECC55', '#88BB44', '#FFBBBB'])
plt.rc('axes', facecolor='#E6E6E6', edgecolor='none',
       axisbelow=True, grid=True, prop_cycle=colors)
plt.rc('grid', color='w', linestyle='solid')
plt.rc('xtick', direction='out', color='gray')
plt.rc('ytick', direction='out', color='gray')
plt.rc('patch', edgecolor='#E6E6E6')
plt.rc('lines', linewidth=2)
plt.rc('figure', figsize=(20, 8))


# print(plt.rcParams.keys())

# plt.rcParams.update({"font.size": 5, "figure.figsize": (20, 8)}, {"xticks.rotation":'vertical'})
xl_file = "D:/00_herne/_export_SmartSheet/IZ224_Weld_Map.xlsx"

df = pd.read_excel(xl_file, sheet_name="IZ224_Weld_Map")
writer = pd.ExcelWriter("D:/00_herne/00_template/WF_temp.xlsx")
# print(df)

#for i in set([c for c in df["Line KKS"]]):
#    print(df["Line KKS"].value_counts())
# df2 = pd.DataFrame(columns=["KKS", "WNO"]).append(df["Line KKS"].value_counts())
# df2 = df.filter([df["Line KKS"].value_counts()], axis=1)
df["counts"] = df.groupby(["Line KKS"])["W_NO"].transform("count")
df_filtered = pd.DataFrame(df.filter(["Line KKS", "counts"], axis=1).drop_duplicates())
# df_filtered.plot(kind="scatter", x="Line KKS", y="counts", title="KKS vs Weld number")
# df_filtered["counts"].plot(kind="hist", title="welds")

# plt.show()

print(df[df["Line KKS"] == "60LBA20BR004"])
# print(df)
# print(df_filtered)

df[df["Line KKS"] == "60LBA20BR004"].to_excel(writer, sheet_name="04_Matrix", startrow=5, header=False, index=False)
writer.save()