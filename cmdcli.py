import cmd
import argparse
from log_config import logger
from parse import *
from utils import *
from main import process_file

class IPCiCmd(cmd.Cmd):
    prompt = '>>> '  # 设置命令行提示符

    def __init__(self, project_dir):
        super().__init__()
        self.project_dir = project_dir

        self.parser = argparse.ArgumentParser(description="处理考勤文件并生成 Excel 文件")
        
        # 定义命令行参数
        self.parser.add_argument('file_path', help="考勤文件的路径")
        self.parser.add_argument('year', type=int, help="年份")
        self.parser.add_argument('month', type=int, help="月份")
        self.parser.add_argument('--tm', type=int, default=3, help="设置时间阈值，单位分钟，默认 3")
        self.parser.add_argument('--debug', action='store_true', help="开启调试模式，传入 --debug 开启调试")    

    def do_process(self, arg):
        """处理解析命令，格式: process <file_path> <year> <month> [--tm 3] [--debug]"""
        # 使用 argparse 解析输入的参数
        try:
            args = self.parser.parse_args(arg.split())
        except SystemExit:
            return  # 如果参数不合法，argparse 会自动处理并打印错误信息

        # 获取参数
        file_path = args.file_path
        year = args.year
        month = args.month
        threshold_minutes = args.tm
        is_debug = args.debug

        # 打印解析结果（用于调试）
        logger.debug(f"文件路径: {file_path}, 年份: {year}, 月份: {month}, 时间阈值: {threshold_minutes}, 调试模式: {is_debug}")
        
        # 调用文件处理函数
        process_file(file_path, year, month, threshold_minutes, is_debug)

    def do_filter(self, arg):
        """处理过滤命令，格式: filter <file_path> <year> <month> [--tm 3] [--debug]"""
        # 使用 argparse 解析输入的参数
        try:
            args = self.parser.parse_args(arg.split())
        except SystemExit:
            return  # 如果参数不合法，argparse 会自动处理并打印错误信息

        # 获取参数
        file_path = args.file_path
        year = args.year
        month = args.month
        threshold_minutes = args.tm
        is_debug = args.debug

        # 打印解析结果（用于调试）
        logger.debug(f"文件路径: {file_path}, 年份: {year}, 月份: {month}, 时间阈值: {threshold_minutes}, 调试模式: {is_debug}")
        
        src_dict = convert_file(file_path)
        filter_dict = filter_times(src_dict, year, month, threshold_minutes, is_debug)
        if(is_debug):
            save_debug_data(filter_dict, "filter")
    
    def do_parse(self, arg):
        """处理解析命令，格式: parse <file_path> [--debug]"""
        # 使用 argparse 解析输入的参数
        try:
            args = self.parser.parse_args(arg.split())
        except SystemExit:
            return  # 如果参数不合法，argparse 会自动处理并打印错误信息

        # 获取参数
        file_path = args.file_path
        is_debug = args.debug
        
        # 打印解析结果（用于调试）
        logger.debug(f"文件路径: {file_path}, 调试模式: {is_debug}")
        
        src_dict = convert_file(file_path)
        if(is_debug):
            save_debug_data(src_dict, "convert")


    def do_exit(self, arg):
        """退出交互式命令行"""
        print("退出交互模式")
        return True

    def do_help(self, arg):
        """显示帮助信息"""
        print("命令:")
        print("  parse <file_path> <year> <month> [-debug] 解析文件并生成结果")
        print("  exit                             退出交互模式")