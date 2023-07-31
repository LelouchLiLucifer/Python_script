import pandas as pd
import numpy as np

def process_data(file_path, output_file_path):
    # 读取Excel文件，并将第一行作为列名
    df = pd.read_excel(file_path, header=0)

    # 查看数据
    print(df.head())

    df['A睡眠质量'] = df.iloc[:, 17] - 1

    print(df.head())

    # 创建新的数据框df_tem
    df_tem = pd.DataFrame()

    # 生成"B1"列的数据（第3列的值减一）
    df_tem['B1'] = df.iloc[:, 2] - 1

    # 生成"B2"列的数据（第8列的值减一）
    df_tem['B2'] = df.iloc[:, 7] - 1

    df_tem['B'] = df_tem[['B1', 'B2']].sum(axis=1)

    # 处理"B"列的数据
    def process_value(value):
        if value == 0:
            return 0
        elif 0 < value < 3:
            return 1
        elif 2 < value < 5:
            return 2
        else:
            return 3

    # 将处理过的数据作为df的一列新数据，列名为"B入睡时间"
    df['B入睡时间'] = df_tem['B'].apply(process_value)

    # 将"17、(2)：___。(不等于卧床时间)"列的各行数据除以60
    df_tem['C'] = df['17、(2)：___。(不等于卧床时间)'] / 60

    # 将结果与"17、(1)近1个月，每夜通常实际睡眠___"列数据求和，并作为新的列"C"
    df_tem['C'] += df['17、(1)近1个月，每夜通常实际睡眠___']

    def process_Cvalue(value):
        if value > 7:
            return 0
        elif 6 < value <= 7:
            return 1
        elif 5 < value <= 6:
            return 2
        else:
            return 3

    # 将处理过的数据作为df的一列新数据，列名为"C睡眠时间"
    df['C睡眠时间'] = df_tem['C'].apply(process_Cvalue)

    # 查看新的数据框df_tem
    print(df_tem.head())

    # 将第一列数据作为上床睡觉的时间的小时
    bedtime_hours = df.iloc[:, 0]

    # 将第二列数据作为上床睡觉的时间的分钟
    bedtime_minutes = df.iloc[:, 1]

    # 将第四列数据作为起床睡觉的时间的小时
    wakeup_hours = df.iloc[:, 3]

    # 将第五列数据作为起床睡觉的时间的分钟
    wakeup_minutes = df.iloc[:, 4]

    # 计算实际睡眠时间
    actual_sleep_time = []

    for i in range(len(df)):
        bedtime = pd.Timestamp(year=2000, month=1, day=1, hour=bedtime_hours[i], minute=bedtime_minutes[i])
        wakeup = pd.Timestamp(year=2000, month=1, day=1, hour=wakeup_hours[i], minute=wakeup_minutes[i])
        
        if bedtime > wakeup:
            # 上床睡觉时间晚于起床时间，说明跨天睡觉
            bedtime -= pd.Timedelta(days=1)
        
        sleep_time = wakeup - bedtime
        actual_sleep_time.append(sleep_time.total_seconds() / 3600)  # Convert sleep time to hours

    # 将实际睡眠时间作为df_tem的一列新数据，列名为"床上时间"
    df_tem['床上时间'] = actual_sleep_time

    df_tem['睡眠效率'] = df_tem['C'] / df_tem['床上时间']

    # 处理"D睡眠效率"列的数据
    def process_sleep_efficiency(value):
        if value >= 0.85:
            return 0
        elif 0.75 <= value < 0.85:
            return 1
        elif 0.65 <= value < 0.75:
            return 2
        else:
            return 3

    # 将处理过的数据作为df的一列新数据，列名为"D睡眠效率"
    df['D睡眠效率'] = df_tem['睡眠效率'].apply(process_sleep_efficiency)

    # 将第9到第17列的各行数据分别求和，并将求和结果分别减去9，然后存储在"E"列中
    df_tem['E'] = df.iloc[:, 8:17].sum(axis=1) - 9

    # 处理"E"列的数据
    def process_sleep_disorder(value):
        if value == 0:
            return 0
        elif 1 <= value <= 9:
            return 1
        elif 10 <= value <= 18:
            return 2
        elif 19 <= value <= 27:
            return 3
        else:
            return None  # Handle other cases, if needed

    # 将处理过的数据作为df的一列新数据，列名为"E.睡眠障碍"
    df['E睡眠障碍'] = df_tem['E'].apply(process_sleep_disorder)

    # 将“20、近1个月，您用药物催眠的情况”列的各行数据均减去1，并作为df的一列新数据，列名为"F催眠药物"
    df['F催眠药物'] = df['20、近1个月，您用药物催眠的情况'] - 1

    df_tem['G'] = df[['21、近1个月，您常感到困倦吗', '22、近1个月，您做事情的精力不足吗']].sum(axis=1) - 2

    # 处理"G"列的数据
    def process_daytime_dysfunction(value):
        if value == 0:
            return 0
        elif 1 <= value <= 2:
            return 1
        elif 3 <= value <= 4:
            return 2
        elif 5 <= value <= 6:
            return 3
        else:
            return None  # Handle other cases, if needed

    # 将处理过的数据作为df的一列新数据，列名为"G日间功能障碍"
    df['G日间功能障碍'] = df_tem['G'].apply(process_daytime_dysfunction)

    # 计算总分列
    df['总分'] = df[['A睡眠质量', 'B入睡时间', 'C睡眠时间', 'D睡眠效率', 'E睡眠障碍', 'F催眠药物', 'G日间功能障碍']].sum(axis=1)

    # 将df输出到指定路径中
    df.to_excel(output_file_path, index=False)
    print(f"处理后的数据已导出到 {output_file_path} 文件。")


def main():
    print("请提供需要处理的Excel文件路径：")

    # 提示用户输入文件路径
    file_path = input("文件路径: ")
    try:
        print("数据处理中...")
        # 提示用户输入输出文件路径
        output_file_path = input("输出文件路径(以output.xlsx结尾): ")
        process_data(file_path, output_file_path)
    except FileNotFoundError:
        print("错误：指定的文件路径不存在。")
    except Exception as e:
        print(f"出现错误：{str(e)}")

if __name__ == "__main__":
    main()