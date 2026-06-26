import pandas as pd

def convert_format(input_csv, output_excel):
    print(f"正在读取 {input_csv} ...")
    df_csv = pd.read_csv(input_csv)
    
    # 创建一个新的 DataFrame，包含目标 Excel 的所有列
    target_columns = ['会员', '级别', '评价星级', '评价内容', '时间', '点赞数', '评论数', '追评时间', '追评内容', '商品属性', '页面标题', '好评度', '评价关键词']
    df_excel = pd.DataFrame(columns=target_columns)
    
    # 字段映射
    df_excel['会员'] = df_csv['用户昵称']
    df_excel['级别'] = df_csv['用户级别']
    df_excel['评价星级'] = df_csv['评分']
    df_excel['评价内容'] = df_csv['评论内容']
    df_excel['时间'] = df_csv['评论时间']
    df_excel['点赞数'] = df_csv['点赞数']
    df_excel['评论数'] = df_csv['回复数']
    
    # 拼接商品属性
    df_excel['商品属性'] = df_csv.apply(lambda row: f"{row.get('商品颜色', '')} {row.get('商品规格', '')}".strip(), axis=1)
    
    df_excel['页面标题'] = df_csv['订单商品名称']
    
    # CSV 中没有的字段，填充为空或默认值
    df_excel['追评时间'] = ""
    df_excel['追评内容'] = ""
    df_excel['好评度'] = ""
    df_excel['评价关键词'] = ""
    
    print(f"转换完成，准备保存为 {output_excel} ...")
    df_excel.to_excel(output_excel, index=False)
    print("保存成功！")

if __name__ == "__main__":
    input_file = "jd_100327958810_comments.csv"
    output_file = "美的_转换后.xlsx"
    
    try:
        convert_format(input_file, output_file)
    except FileNotFoundError:
        print(f"错误: 找不到文件 '{input_file}'，请确认它是否存在。")
