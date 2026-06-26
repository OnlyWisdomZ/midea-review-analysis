import pandas as pd

def clean_excel_format(input_excel, output_excel):
    print(f"正在读取 {input_excel} ...")
    # 读取未清洗的 Excel 数据
    df_raw = pd.read_excel(input_excel)
    
    # 目标 美的.xlsx 的所需列结构
    target_columns = [
        '会员', '级别', '评价星级', '评价内容', '时间', '点赞数', 
        '评论数', '追评时间', '追评内容', '商品属性', '页面标题', 
        '好评度', '评价关键词'
    ]
    
    # 过滤掉不需要的列（如 页面网址, sku, 评论类型, 该类型评论数），只保留目标列
    # 对于 target_columns 中可能在 df_raw 中缺失的列，使用 reindex 可以安全地保留并用 NaN 填充
    df_clean = df_raw.reindex(columns=target_columns)
    
    print(f"提取完成，包含 {len(df_clean.columns)} 列。准备保存为 {output_excel} ...")
    
    # 保存为新的 Excel 文件
    df_clean.to_excel(output_excel, index=False)
    print("保存成功！")

if __name__ == "__main__":
    input_file = "美的-未清洗.xlsx"
    output_file = "美的_清洗后.xlsx"
    
    try:
        clean_excel_format(input_file, output_file)
    except FileNotFoundError:
        print(f"错误: 找不到文件 '{input_file}'，请确认它是否存在。")
    except Exception as e:
        print(f"处理过程中发生错误: {e}")
