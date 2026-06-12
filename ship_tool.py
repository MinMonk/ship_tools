#!/usr/bin/env python3
import argparse
import logging
import os
import sys

from src.utils.common_utils import setup_logging, load_basic_data
from src.fba_label.process_fba_label import process_fba_label
from src.invoice.process_ship_invoice import read_invoice_mode, run_invoice, shipment_input_dir, validate_shipment_input_dir
from src.plan.process_ship_plan import run_plan
from src.price.process_ship_price import run_calc
from src.summary.process_summary import run_summary

LABEL_MODES = {"FIST", "SEND"}


def validate_label_input_dir(input_dir):
    if not os.path.isdir(input_dir):
        sys.exit(f"物流模式对应的标签输入目录不存在：{input_dir}")

    pdf_files = [name for name in os.listdir(input_dir) if name.lower().endswith(".pdf")]
    if not pdf_files:
        sys.exit(f"输入目录 {input_dir} 下未找到外箱标签 PDF 文件。")

    return {"pdf_count": len(pdf_files)}


def run_label(mode, input_dir, logger, log_input=True):
    if log_input:
        logger.info(f"\033[36m▶ 物流模式：{mode}，输入目录：{input_dir}\033[0m")

    logger.info("\033[33m ----------------------------------------------- \033[0m")

    logger.info(f"\033[36m▶ 裁剪【FBA外箱标】开始执行  \033[0m")
    process_fba_label(mode=mode, label_type="fba", input_dir=input_dir)
    logger.info(f"\033[36m▶ 裁剪【FBA外箱标】执行完毕  \033[0m")

    if mode == 'SEND':
        logger.info("\033[33m ----------------------------------------------- \033[0m")

        logger.info(f"\033[36m▶ 裁剪【唛头】开始执行  \033[0m")
        process_fba_label(mode=mode, label_type="mark", input_dir=input_dir)
        logger.info(f"\033[36m▶ 裁剪【唛头】执行完毕  \033[0m")


def main():
    parser = argparse.ArgumentParser(
        prog='ship_tool.py',
        description=(
            'Ship Tool CLI\n'
            '可选指令:\n'
            '  invoice: 根据亚马逊后台货件装箱单数据，结合 plan_data/开票信息.xlsx，\n'
            '           生成货代公司的交接资料(发票)，并按 G1 物流模式处理标签\n'
            '  plan   : 根据 plan_data/出货计划.xlsx，创建亚马逊后台批量货件模板及\n'
            '           领星 ERP 出货计划模板，方便导入系统\n'
            '  summary: 汇总 shipment_data 下全部装箱单的箱数、盒数、SKU、体积和重量\n'
            '  calc   : 根据 summary 重量和 shipment_data/ship_price.xlsx 计算海运费\n'
            '  label  : 根据指定 FIST/SEND 模式，独立裁剪 shipment_data 下的标签 PDF'
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例:\n"
            "  python3 ship_tool.py invoice\n"
            "  python3 ship_tool.py label SEND"
        )
    )
    parser.add_argument(
        'command',
        choices=['invoice', 'plan', 'summary', 'calc', 'label'],
        help='指定要执行的功能: invoice、plan、summary、calc 或 label'
    )
    parser.add_argument(
        'command_arg',
        nargs='?',
        type=str.upper,
        help='plan 命令可选 US/CA；label 命令必填 FIST/SEND'
    )
    args = parser.parse_args()

    if args.command == 'plan':
        if args.command_arg and args.command_arg not in {'US', 'CA'}:
            parser.error('plan 命令可选参数仅支持 US 或 CA')
        plan_country = args.command_arg
    elif args.command == 'label':
        if not args.command_arg:
            parser.error('label 命令必须指定 FIST 或 SEND')
        if args.command_arg not in LABEL_MODES:
            parser.error('label 命令参数仅支持 FIST 或 SEND')
        label_mode = args.command_arg
        plan_country = None
    else:
        if args.command_arg:
            parser.error(f'{args.command} 命令不支持额外参数')
        plan_country = None

    # 初始化日志
    setup_logging()
    logger = logging.getLogger(__name__)

    if args.command == 'invoice':
        invoice_mode = read_invoice_mode()
        shipment_data_dir = shipment_input_dir(invoice_mode)
        validate_shipment_input_dir(shipment_data_dir)
        logger.info(f"\033[36m▶ 物流模式：{invoice_mode}，输入目录：{shipment_data_dir}\033[0m")
        # 加载 basic_data 下的全局缓存
        global_info, product_dict, package_dict = load_basic_data()
        logger.info(f"\033[36m▶ INVOICE 开始执行  \033[0m")
        run_invoice(global_info, product_dict, package_dict, shipment_data_dir)
        logger.info(f"\033[36m▶ INVOICE 执行完毕  \033[0m")

        run_label(invoice_mode, shipment_data_dir, logger, log_input=False)

    elif args.command == 'plan':
        # 加载 basic_data 下的全局缓存
        global_info, product_dict, package_dict = load_basic_data()
        logger.info(f"\033[36m▶ PLAN 开始执行  \033[0m")
        run_plan(global_info, package_dict, plan_country)
        logger.info(f"\033[36m▶ PLAN 执行完毕  \033[0m")
    elif args.command == 'summary':
        # 加载 basic_data 下的全局缓存
        global_info, product_dict, package_dict = load_basic_data()
        logger.info(f"\033[36m▶ SUMMARY 开始执行  \033[0m")
        run_summary(product_dict, package_dict)
        logger.info(f"\033[36m▶ SUMMARY 执行完毕  \033[0m")
    elif args.command == 'calc':
        # 加载 basic_data 下的全局缓存
        global_info, product_dict, package_dict = load_basic_data()
        logger.info(f"\033[36m▶ CALC 开始执行  \033[0m")
        run_calc(product_dict, package_dict)
        logger.info(f"\033[36m▶ CALC 执行完毕  \033[0m")
    elif args.command == 'label':
        shipment_data_dir = shipment_input_dir(label_mode)
        validate_label_input_dir(shipment_data_dir)
        logger.info(f"\033[36m▶ LABEL 开始执行  \033[0m")
        run_label(label_mode, shipment_data_dir, logger)
        logger.info(f"\033[36m▶ LABEL 执行完毕  \033[0m")
    else:
        logger.error(f'未知命令: {args.command}')

if __name__ == '__main__':
    main()
