import logging
from typing import Dict, Any


logger = logging.getLogger(__name__)


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


def run_summary(product_dict: Dict[str, Dict], package_dict: Dict[str, Dict]) -> Dict[str, Dict[str, Any]]:
    """
    summary 命令入口：读取 shipment_data 下全部装箱单 CSV 并输出仓库维度汇总。
    """
    from src.invoice.process_ship_invoice import parse_shipment_csv

    shipment_data = parse_shipment_csv("./shipment_data/", product_dict)
    summaries = {}
    for shipment_id, ship_data in shipment_data.items():
        repo_name = (ship_data.get("shipment_info") or {}).get("仓库名称")
        summary = summarize_shipment(ship_data, package_dict)
        summaries[shipment_id] = summary
        logger.info(f"【{shipment_id}】-【{repo_name}】: {summary}")

    return summaries
