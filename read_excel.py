import pandas as pd

# 读取Excel文件
file_path = '2026年4月-安恒开通服务情况统计表V6.0.xlsx'

# 获取所有工作表名称
xls = pd.ExcelFile(file_path)
sheet_names = xls.sheet_names

print("=== Excel文件结构分析 ===")
print(f"文件名: {file_path}")
print(f"工作表数量: {len(sheet_names)}")
print(f"工作表名称: {sheet_names}\n")

# 遍历每个工作表
for sheet_name in sheet_names:
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    print(f"--- 工作表: {sheet_name} ---")
    print(f"行数: {len(df)}")
    print(f"列数: {len(df.columns)}")
    print(f"列名: {list(df.columns)}")
    print("\n前5行数据预览:")
    print(df.head())
    print("\n" + "="*60 + "\n")

# 将数据保存为CSV以便后续处理
for sheet_name in sheet_names:
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    csv_filename = f"{sheet_name}.csv"
    df.to_csv(csv_filename, index=False, encoding='utf-8-sig')
    print(f"已保存CSV文件: {csv_filename}")
