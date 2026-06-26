import pandas as pd

def handle_missing_values(input_file, output_file):
    print(f"Loading data from {input_file}...")
    try:
        # 读取Excel文件
        df = pd.read_excel(input_file)
        initial_count = len(df)
        print(f"原始数据行数: {initial_count}")
        # 允许包含缺失值的列名（不对这些列的缺失值进行行删除）
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
        # 只有当这些列中存在缺失值时，才会删除对应的行
        cols_to_check = [col for col in df.columns if col not in allowed_missing_cols]
        print(f"需要检查缺失值的列数: {len(cols_to_check)} / {len(df.columns)}")
        # 剔除在 cols_to_check 中含有任何缺失值的行
        df_cleaned = df.dropna(subset=cols_to_check)
        final_count = len(df_cleaned)
        print(f"处理后数据行数: {final_count}")
        print(f"共删除 {initial_count - final_count} 行含有缺失值的数据。")
        # 保存到新文件
        df_cleaned.to_excel(output_file, index=False)
        print(f"处理缺失值后的数据已保存至: {output_file}")
    except Exception as e:
        print(f"处理文件时出错: {e}")
if __name__ == "__main__":
    input_filename = "美的_已去重.xlsx"  # 建议使用去重后的文件作为输入，也可以改为 "美的.xlsx"
    output_filename = "美的_已清洗.xlsx"
    handle_missing_values(input_filename, output_filename)
