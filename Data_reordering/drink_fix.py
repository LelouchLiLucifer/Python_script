import pandas as pd

def process_data(column_data):
    # 将指定列的数据按分隔符"┋"分隔，并转换为一个新的DataFrame
    processed_data = column_data.str.split('┋', expand=True)
    
    # 创建一个新的DataFrame dftem，将处理后的数据放入其中
    dftem = pd.DataFrame(processed_data.values, columns=[f"第11列_分隔_{i+1}" for i in range(processed_data.shape[1])])
    
    return dftem

if __name__ == "__main__":
    file_name = "drink.xlsx"
    try:
        # 读取Excel文件并将第一行作为列名
        df = pd.read_excel(file_name, header=0)
        
        column_number = 5  # 处理第6列数据
        column_data = df.iloc[:, column_number]
        
        dftem = process_data(column_data)
        
        print("成功导入Excel文件并生成数据框 dftem:")
        print(dftem)
        
    except FileNotFoundError:
        print(f"文件 '{file_name}' 未找到。请确保文件名和路径正确。")
    except Exception as e:
        print(f"出现错误：{e}")

df0 = pd.DataFrame(columns=['A', 'B', 'C', 'D', 'E'])

# 处理每行数据的函数
def extract_data(row):
    drink_categories = {'A': '洋酒（威士忌，伏特加，龙舌兰，白兰地，朗姆等）',
                        'B': '啤酒',
                        'C': '红酒',
                        'D': '酿造粮酒（中国白酒）',
                        'E': '鸡尾酒（配制酒）'}
    data = {}
    for category, keyword in drink_categories.items():
        for col in row.index:
            if isinstance(row[col], str) and keyword in row[col]:
                data[category] = row[col]
                break
        else:
            data[category] = pd.NA
    return pd.Series(data)

# 将函数应用到每一行，并创建一个新的DataFrame
df0 = dftem.apply(extract_data, axis=1)

print(df0)

output_file_path = "output.xlsx"
df0.to_excel(output_file_path, index=False)