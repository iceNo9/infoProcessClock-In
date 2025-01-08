import datetime
import re
import os
import sys
import subprocess
import openpyxl
import argparse
from collections import defaultdict
from attendanceManager import AttendanceManager

# 设置最大时间差
TIME_DIFF_THRESHOLD = 3  # 时间差阈值（单位：分钟）

# 获取中文星期
def get_weekday_chinese(date_str):
    date = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
    weekday_num = date.weekday()  # 返回0=Monday, 1=Tuesday, ..., 6=Sunday
    weekday_chinese = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
    return weekday_chinese[weekday_num]

# 生成并保存 Excel 文件
def generate_excel_file(wb, final_attendance_data):
    # 删除默认的工作表
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']

    # 创建一个工作表名为 "src"
    ws = wb.create_sheet(title="src")

    # 写入表头
    ws.append(["日期", "星期", "打卡时间1", "打卡时间2", "打卡时间3", "打卡时间4"])  # 可根据最大打卡次数调整

    # 遍历最终考勤数据
    sorted_dates = sorted(final_attendance_data.keys())
    for date in sorted_dates:
        weekday = get_weekday_chinese(date)
        formatted_times = [time.split(" ")[1] for time in final_attendance_data[date]]  # 取出时分秒部分

        # 将同一天的打卡时间写入同一行，每个时间占据一个单元格
        row = [date, weekday] + formatted_times

        # 补齐空白单元格以保证每一行的列数一致（假设最多有5次打卡）
        while len(row) < 7:  # 确保每行至少有6列（日期、星期及5个打卡时间）
            row.append('')

        ws.append(row)

# 读取文件并处理考勤数据
def process_attendance_file(file_path, year, month):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
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

            # 按日期处理并比较时间差
            final_attendance_data = {}

            for date, times in attendance_data.items():
                # 将字符串时间转换为 datetime 对象，确保包含年月日
                times = [datetime.datetime.strptime(datetime_str, "%Y-%m-%d %H:%M:%S") for datetime_str in times]
                times.sort()  # 按时间排序

                merged_times = []
                current_time = times[0]

                for next_time in times[1:]:
                    time_diff = (next_time - current_time).total_seconds() / 60  # 计算时间差（分钟）

                    if time_diff <= TIME_DIFF_THRESHOLD:
                        # 如果时间差小于等于阈值，选择较大的时间
                        current_time = max(current_time, next_time)
                    else:
                        # 如果时间差大于阈值，保存当前时间，并更新当前时间为 next_time
                        merged_times.append(current_time)
                        current_time = next_time

                # 添加最后一个时间点
                merged_times.append(current_time)

                # 存储结果，保留完整的年月日时分秒
                final_attendance_data[date] = [time.strftime("%Y-%m-%d %H:%M:%S") for time in merged_times]

            # 创建 Excel 工作簿
            wb = openpyxl.Workbook()

            # 生成 Excel 文件
            generate_excel_file(wb, final_attendance_data)

            # 设置输出文件路径
            output_file = file_path.replace(".txt", ".xlsx")
            
            # 假设 attendance_manager 是一个有效的对象，并调用它来处理考勤数据
            attendance_manager = AttendanceManager(year, is_flexible=True)
            result = attendance_manager.process_month(month, final_attendance_data)
            attendance_manager.write_attendance_to_excel(wb, result)

            wb.save(output_file)
            print(f"数据已保存为：{output_file}")

    except FileNotFoundError:
        print("文件未找到")
    except Exception as e:
        print(f"发生错误: {e}")

# 主函数，传入文件路径
if __name__ == "__main__":    
   
    # 这里输出最新的commit提交内容
    # 获取最新的 Git 提交内容
    commit_message = ""
    

    # 判断是否是 Nuitka 打包环境
    if getattr(sys, 'frozen', False):
        print("运行环境: 打包")
        
        # 设置命令行参数解析器
        parser = argparse.ArgumentParser(description="处理考勤文件并生成 Excel 文件")
        parser.add_argument('file_path', help="考勤文件的路径")
        parser.add_argument('year', type=int, help="年份")
        parser.add_argument('month', type=int, help="月份")

        args = parser.parse_args()

        input_file_path = args.file_path  # 获取传入的文件路径
        year = args.year  # 获取传入的年份
        month = args.month  # 获取传入的月份

        commit_message="feat:初版作成，支持2024.12--2025.2考勤识别；"
    else:
        # 在 Nuitka 打包环境下使用自定义的数据
        input_file_path = "test.txt"  
        year = 2025  # 使用预设年份
        month = 1  # 使用预设月份

        try:
            # 获取最近一次提交的 commit message
            commit_message = subprocess.check_output(["git", "log", "-1", "--pretty=%B"]).strip().decode("utf-8")
        except subprocess.CalledProcessError as e:
            commit_message = "无法获取最新的 commit 信息"

        print("运行环境: 调试")

    VER = "V1.0-"
    DATE = "20250108"    

    print(f"版本: {VER}{DATE}")
    print("版本修改说明: " + commit_message)

    # 调用处理考勤文件函数
    process_attendance_file(input_file_path, year, month)
