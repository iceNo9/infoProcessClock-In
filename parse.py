from collections import defaultdict
import re
from log_config import logger
from datetime import datetime, timedelta
import chardet

def convert_file(file_path):
    try:
        # 自动识别文件编码
        with open(file_path, 'rb') as file:
            raw_data = file.read()
            result = chardet.detect(raw_data)
            encoding = result['encoding']
        
        # 使用识别的编码打开文件
        with open(file_path, 'r', encoding=encoding) as file:
            content = file.read()

            # 使用正则表达式提取所有日期和时间
            pattern = r'(\d{4}-\d{2}-\d{2}) (\d{2}:\d{2}:\d{2})'
            matches = re.findall(pattern, content)

            # 使用 defaultdict 来存储数据
            attendance_data = defaultdict(list)

            # 将匹配的日期和时间存储到字典中，确保时间包含年月日
            for date, time in matches:
                datetime_str = f"{date} {time}"
                attendance_data[date].append(datetime_str)

            return attendance_data

    except Exception as e:
        logger.error(f"解析出错,文件:{file_path},异常:{e}")
    
def filter_times(input_dict, year, month, threshold_minutes=3):
    # 设置时间阈值
    threshold = timedelta(minutes=threshold_minutes)
    
    # 过滤后的字典
    filtered_dict = {}

    # 遍历字典
    for date, times in input_dict.items():
        # 将日期字符串转换为 datetime 对象
        date_obj = datetime.strptime(date, "%Y-%m-%d")
        
        # 只保留符合指定年份和月份的日期
        if date_obj.year != year or date_obj.month != month:
            continue
        
        # 转换时间字符串为 datetime 对象并排序
        time_objects = [datetime.strptime(time, "%Y-%m-%d %H:%M:%S") for time in times]
        time_objects.sort()  # 按时间升序排列
        
        # 用于存储过滤后的时间
        filtered_times = []

        for time in time_objects:
            if not filtered_times:  # 第一个时间直接加入
                filtered_times.append(time)
            else:
                # 获取当前时间与上一个保留时间的差值
                if time - filtered_times[-1] <= threshold:
                    # 如果差值小于等于阈值，保留较大的时间
                    # 当前时间小于等于阈值，更新为当前时间中较大的时间
                    filtered_times[-1] = max(filtered_times[-1], time)
                else:
                    # 否则，保留当前时间
                    filtered_times.append(time)

        # 转换回字符串格式并保存到新字典
        filtered_dict[date] = [time.strftime("%Y-%m-%d %H:%M:%S") for time in filtered_times]

    return filtered_dict