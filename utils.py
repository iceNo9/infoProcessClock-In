from log_config import logger
import json
import os

def save_debug_data(attendance_data, dir, name):
    try:
        # 设置 debug 文件夹路径
        debug_dir = os.path.join(dir, 'debug')

        # 确保 debug 文件夹存在
        if not os.path.exists(debug_dir):
            os.makedirs(debug_dir)

        # 设置输出的 json 文件路径
        json_file_path = os.path.join(debug_dir, f"{name}.json")

        # 将字典转换为 JSON 格式并写入文件
        with open(json_file_path, 'w', encoding='utf-8') as json_file:
            json.dump(attendance_data, json_file, ensure_ascii=False, indent=4)

        logger.debug(f"Debug data saved to: {json_file_path}")

    except Exception as e:
        logger.error(f"保存调试数据出错,异常:{e}")