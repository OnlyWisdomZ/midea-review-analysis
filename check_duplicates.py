import pandas as pd

def check_duplicates(file_path):
    print(f"Loading data from {file_path}...")
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)
        total_rows = len(df)
        
        # 查找重复行 (keep=False表示标记所有重复的行，包括第一次出现的)
        # 如果只想知道有多少组重复的，可以使用 keep='first'
        duplicates = df[df.duplicated(keep=False)]
        num_duplicates = len(duplicates)
        
        # 计算完全重复的行数（基于 keep='first'）
        # 这样能准确反映还有多少条冗余数据
        redundant_rows = df.duplicated().sum()
        
        print(f"总数据行数: {total_rows}")
        print(f"发现完全相同的冗余记录条数: {redundant_rows}")
        
        if redundant_rows > 0:
            print(f"警告：文件中仍存在 {redundant_rows} 条冗余的重复数据。")
            print("以下是部分重复数据示例：")
            # 打印包含所有重复项的数据（按原顺序）
            print(duplicates.sort_values(by=list(df.columns)).head(10))
            
            # 如果需要，可以将重复项保存到新文件以供检查
            # duplicates.sort_values(by=list(df.columns)).to_excel("重复项检查.xlsx", index=False)
        else:
            print("恭喜！数据中没有发现任何完全重复的行，去重结果完美。")
            
    except Exception as e:
        print(f"检查重复值时出错: {e}")

if __name__ == "__main__":
    file_name = "美的_已去重.xlsx"
    check_duplicates(file_name)
