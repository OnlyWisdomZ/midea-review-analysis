import pandas as pd

def check_missing_values(file_path):
    print(f"Loading data from {file_path}...")
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)
        total_rows = len(df)
        print(f"总数据行数: {total_rows}")
        
        # 允许包含缺失值的列名
        allowed_missing_cols = [
            '级别', 
            '点赞数', 
            '评论数', 
            '追评时间', 
            '追评内容', 
            '好评度', 
            '评价关键词'
        ]
        
        # 找出需要进行缺失值检查的列（所有列减去允许有缺失值的列）
        cols_to_check = [col for col in df.columns if col not in allowed_missing_cols]
        print(f"正在检查 {len(cols_to_check)} 个不允许有缺失值的列...")
        
        # 检查这些列中是否有缺失值
        missing_info = df[cols_to_check].isnull().sum()
        cols_with_missing = missing_info[missing_info > 0]
        
        if not cols_with_missing.empty:
            print("警告：在不允许有缺失值的列中发现了缺失值！")
            for col, count in cols_with_missing.items():
                print(f"  - 列 '{col}' 包含 {count} 个缺失值。")
        else:
            print("恭喜！所有不允许有缺失值的列均没有缺失值，清洗结果完美。")
            
        # 打印允许缺失的列的缺失情况（作为参考）
        print("\n--- 补充信息：允许有缺失值的列目前的缺失情况 ---")
        cols_in_df = [col for col in allowed_missing_cols if col in df.columns]
        if cols_in_df:
            allowed_missing_info = df[cols_in_df].isnull().sum()
            for col, count in allowed_missing_info.items():
                print(f"  - 列 '{col}' 包含 {count} 个缺失值 (允许范围内)。")
                
    except Exception as e:
        print(f"检查缺失值时出错: {e}")

if __name__ == "__main__":
    file_name = "美的_已清洗.xlsx"
    check_missing_values(file_name)
