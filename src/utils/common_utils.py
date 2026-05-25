import logging
import csv
import pandas as pd
from datetime import datetime

def setup_logging():
    """
    全局日志配置：打印时间、级别、消息
    """
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s %(filename)-20s:%(lineno)4d - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

def parse_txt_to_dict(file_path):
    """解析给定文件中的键值对，并返回一个字典。"""
    parsed_dict = {}
    with open(file_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            # 如果该行为空，直接跳过
            if not line:
                continue

            # 以第一个 '=' 为分隔符进行拆分，避免 value 中带有 '=' 的情况被错误切割
            key, value = line.split('=', 1)

            # 存入字典
            parsed_dict[key.strip()] = value.strip()

    return parsed_dict

def parse_csv_to_dict(file_path, key_field):
    """
    通用 CSV 解析函数：指定文件路径和 Key 列名，
    返回一个以 key_field 内容为键，对应整行数据(字典形式)为值的字典。
    :param file_path: CSV 文件路径
    :param key_field: 作为字典 key 的列名（该列需在 CSV 表头中存在）
    :return: { key_field_value: {列名: 对应值, ...}, ... }
    """
    result = {}
    # 使用 'utf-8-sig' 编码可以避免 BOM 头导致字段名解析异常
    with open(file_path, mode='r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            # 如果 key_field 不在表头中，抛出错误
            if key_field not in row:
                raise ValueError(f"指定的 Key 列名 `{key_field}` 不存在于 CSV 表头中。")
            key_value = row[key_field]
            result[key_value] = row
    return result

def parse_csv_to_dict2(file_path, key_field):
    """
    通用 CSV 解析函数：指定文件路径和 Key 列名或列名列表，
    返回一个以 key_field 内容为键（单个字段或多字段拼接），
    对应整行数据(字典形式)为值的字典。
    :param file_path: CSV 文件路径
    :param key_field: 作为字典 key 的列名（字符串）或列名列表/元组
    :return: { key_value: {列名: 对应值, ...}, ... }
    :raises ValueError: 指定的字段在表头中不存在
    """
    result = {}
    # 使用 'utf-8-sig' 编码可以避免 BOM 头导致字段名解析异常
    with open(file_path, mode='r', encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        headers = reader.fieldnames or []
        
        # 规范为列表，便于后续处理
        if isinstance(key_field, str):
            keys = [key_field]
        else:
            keys = list(key_field)
        
        # 校验所有 key 字段都在表头中
        missing = [k for k in keys if k not in headers]
        if missing:
            raise ValueError(f"指定的 Key 列名不存在于 CSV 表头中: {missing}")
        
        for row in reader:
            # 拼接 key 值
            key_value = "-".join(str(row[k]).strip() for k in keys)
            result[key_value] = row

    return result

def load_basic_data():
    try:
        global_info = parse_txt_to_dict('./basic_data/data.txt')
        product_dict = parse_csv_to_dict('./basic_data/产品信息.csv', 'ASIN')
        package_dict = parse_csv_to_dict('./basic_data/产品包装信息.csv', '套装类型')
        print('[基础数据]加载完成...')
    except FileNotFoundError:
        print("解析[基础数据]失败，根据报错信息排查问题")
        return
    except ValueError as ve:
        print(f"解析[基础数据]出错：{ve}")
        return
    
    return global_info, product_dict, package_dict

def format_current_date(fmt: str = "yyyyMMdd") -> str:
    """
    返回当前日期的格式化字符串

    :param fmt: 日期格式（可选）
        - 默认 "yyyyMMdd"
        - 支持自定义格式，如 "yyyy-MM-dd"、"yyyy/MM/dd HH:mm:ss"
    :return: 格式化后的日期字符串
    """
    now = datetime.now()

    # 将自定义 yyyy/MM/dd 风格映射为 Python strftime 格式
    fmt_map = {
        "yyyy": "%Y",
        "MM": "%m",
        "dd": "%d",
        "HH": "%H",
        "mm": "%M",
        "ss": "%S",
    }

    py_fmt = fmt
    for k, v in fmt_map.items():
        py_fmt = py_fmt.replace(k, v)

    return now.strftime(py_fmt)


def root_dir(type):
    date_str = format_current_date()
    return f"./output/{date_str}/{type}"