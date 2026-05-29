import logging
import csv
from pathlib import Path
from typing import Dict, Any, Optional


logger = logging.getLogger(__name__)

SUMMARY_MODES = ("FIST", "SEND")
SHIPMENT_DATA_ROOT = Path("./shipment_data")


def summarize_shipment(ship_data: Dict[str, Any], package_dict: Dict[str, Dict]) -> Dict[str, Any]:
    """
    汇总单个货件的 SKU 数、盒数、箱数、体积和重量。
    """
    if not ship_data:
        return {
            "总sku数": 0,
            "总盒数": 0,
            "总箱数": 0,
            "总体积(m³)": 0.0,
            "总重量(kg)": 0.0,
        }

    shipment_info = ship_data.get("shipment_info") or {}
    list_data = ship_data.get("list_data") or {}

    sku_set = ship_data.get("sku_set") or set()
    if sku_set:
        sku_count = len(sku_set)
    else:
        sku_count = int(shipment_info.get("SKU 数量", 0) or len(list_data))

    total_qty = 0
    total_boxes = 0
    total_volume_m3 = 0.0
    total_weight_kg = 0.0

    for product_type, data in list_data.items():
        qty = int(data.get("total_num", 0) or 0)
        boxes = int(data.get("total_boxes", 0) or 0)
        total_qty += qty
        total_boxes += boxes

        package_data = package_dict.get(product_type) or {}
        length_cm = float(package_data.get("长(cm)", 0) or 0)
        width_cm = float(package_data.get("宽(cm)", 0) or 0)
        height_cm = float(package_data.get("高(cm)", 0) or 0)
        gross_weight_kg = float(package_data.get("毛重(kg)", 0) or 0)

        total_volume_m3 += length_cm * width_cm * height_cm * boxes / 1_000_000
        total_weight_kg += gross_weight_kg * boxes

    return {
        "总sku数": sku_count,
        "总盒数": total_qty,
        "总箱数": total_boxes,
        "总体积(m³)": round(total_volume_m3, 4),
        "总重量(kg)": round(total_weight_kg, 2),
    }


def parse_summary_csv_file(csv_path: Path, product_dict: Dict[str, Dict], mode: str) -> Optional[Dict[str, Any]]:
    try:
        with csv_path.open("r", encoding="utf-8-sig") as f:
            rows = list(csv.reader(f))
    except Exception as exc:
        logger.info(f"读取文件 {csv_path.name} 失败：{exc}")
        return None

    if len(rows) < 10:
        logger.info(f"文件 {csv_path.name} 行数不足 10 行，跳过。")
        return None

    shipment_info = {}
    for row in rows[:9]:
        if len(row) < 2:
            continue
        key = row[0].strip()
        value = row[1].strip()
        if key:
            shipment_info[key] = value

    if mode == "SEND":
        warehouse_name = csv_path.stem
    else:
        warehouse_name = shipment_info.get("仓库名称") or shipment_info.get("货件名称", "").split("-")[-1] or csv_path.stem
    shipment_info["仓库名称"] = warehouse_name

    shipment_number = shipment_info.get("货件编号") or csv_path.stem
    list_data = {}
    sku_set = set()

    for row in rows[10:]:
        if len(row) < 15:
            continue
        sku = row[0].strip()
        asin = row[2].strip()
        if sku:
            sku_set.add(sku)

        product_type = product_dict.get(asin, {}).get("套装类型", "Unknown")
        if product_type == "Unknown":
            raise ValueError(f"找不到{asin}的包装信息，请检查...")

        try:
            quantity = int(row[13].strip())
        except ValueError:
            quantity = 0
        try:
            boxes = int(row[12].strip())
        except ValueError:
            boxes = 0

        package_seq_num = row[14].strip().replace(shipment_number, "")
        if product_type not in list_data:
            list_data[product_type] = {"total_num": 0, "total_boxes": 0, "package_seq": []}
        list_data[product_type]["total_num"] += quantity
        list_data[product_type]["total_boxes"] += boxes
        if package_seq_num:
            list_data[product_type]["package_seq"].append(package_seq_num)

    return {
        "shipment_info": shipment_info,
        "list_data": list_data,
        "sku_set": sku_set,
    }


def load_summary_shipments(mode: str, product_dict: Dict[str, Dict]) -> Dict[str, Dict[str, Any]]:
    input_dir = SHIPMENT_DATA_ROOT / mode
    if not input_dir.is_dir():
        return {}

    csv_files = sorted(path for path in input_dir.iterdir() if path.is_file() and path.suffix.lower() == ".csv")
    if not csv_files:
        return {}

    shipments = {}
    for csv_path in csv_files:
        ship_data = parse_summary_csv_file(csv_path, product_dict, mode)
        if not ship_data:
            continue
        warehouse_name = (ship_data.get("shipment_info") or {}).get("仓库名称") or csv_path.stem
        shipments[f"{mode}-{warehouse_name}"] = ship_data
    return shipments


def run_summary(product_dict: Dict[str, Dict], package_dict: Dict[str, Dict]) -> Dict[str, Dict[str, Any]]:
    """
    summary 命令入口：分别读取 shipment_data/FIST 和 shipment_data/SEND 下的装箱单 CSV。
    """
    summaries = {}
    for mode in SUMMARY_MODES:
        logger.info(f"▶ {mode} 汇总")
        shipment_data = load_summary_shipments(mode, product_dict)
        if not shipment_data:
            logger.info("暂无数据")
            continue

        for summary_key, ship_data in shipment_data.items():
            summary = summarize_shipment(ship_data, package_dict)
            summaries[summary_key] = summary
            logger.info(f"【{summary_key}】: {summary}")

    return summaries
