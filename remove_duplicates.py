import pandas as pd

def remove_duplicates(input_file, output_file):
    print(f"Loading data from {input_file}...")
    try:
        # 读取Excel文件
        df = pd.read_excel(input_file)
        initial_count = len(df)
        # 去重，默认对所有列进行严格去重
        df_dedup = df.drop_duplicates()
        final_count = len(df_dedup)
        print(f"原始数据行数: {initial_count}")
        print(f"去重后数据行数: {final_count}")
        print(f"共删除 {initial_count - final_count} 行重复数据。")
        # 保存到新文件
        df_dedup.to_excel(output_file, index=False)
        print(f"去重后的数据已保存至: {output_file}")
    except Exception as e:
        print(f"处理文件时出错: {e}")
if __name__ == "__main__":
    input_filename = "美的.xlsx"
    output_filename = "美的_已去重.xlsx"
    remove_duplicates(input_filename, output_filename)
