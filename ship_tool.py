#!/usr/bin/env python3
import argparse
import logging

from src.utils.common_utils import setup_logging, load_basic_data
from src.fba_label.process_fba_label import process_fba_label
from src.invoice.process_ship_invoice import run_invoice
from src.plan.process_ship_plan import run_plan

def main():
    parser = argparse.ArgumentParser(
        prog='ship_tool.py',
        description=(
            'Ship Tool CLI\n'
            '可选指令:\n'
            '  invoice: 根据亚马逊后台货件装箱单数据，结合 plan_data/开票信息.xlsx，\n'
            '           生成货代公司的交接资料(发票)。可选标签处理模式: FIST 或 SEND\n'
            '  plan   : 根据 plan_data/出货计划.xlsx，创建亚马逊后台批量货件模板及\n'
            '           领星 ERP 出货计划模板，方便导入系统'
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例:\n"
            "  python3 ship_tool.py invoice\n"
            "  python3 ship_tool.py invoice SEND"
        )
    )
    parser.add_argument(
        'command',
        choices=['invoice', 'plan'],
        help='指定要执行的功能: invoice 或 plan'
    )
    parser.add_argument(
        'invoice_mode',
        nargs='?',
        default='FIST',
        type=str.upper,
        choices=['FIST', 'SEND'],
        help='invoice 命令的标签处理模式，默认 FIST；SEND 会按奇偶页拆分外箱标签'
    )
    args = parser.parse_args()

    if args.command == 'plan' and args.invoice_mode != 'FIST':
        parser.error('invoice_mode 参数仅支持 invoice 命令')

    # 初始化日志
    setup_logging()
    logger = logging.getLogger(__name__)

    # 加载 basic_data 下的全局缓存
    global_info, product_dict, package_dict = load_basic_data()

    if args.command == 'invoice':
        logger.info(f"\033[36m▶ INVOICE 开始执行  \033[0m")
        run_invoice(global_info, product_dict, package_dict)
        logger.info(f"\033[36m▶ INVOICE 执行完毕  \033[0m")

        logger.info("\033[33m ----------------------------------------------- \033[0m")

        logger.info(f"\033[36m▶ 裁剪FBA外箱标 开始执行  \033[0m")
        process_fba_label(mode=args.invoice_mode)
        logger.info(f"\033[36m▶ 裁剪FBA外箱标 执行完毕  \033[0m")

    elif args.command == 'plan':
        logger.info(f"\033[36m▶ PLAN 开始执行  \033[0m")
        run_plan(global_info, package_dict)
        logger.info(f"\033[36m▶ PLAN 执行完毕  \033[0m")
    else:
        logger.error(f'未知命令: {args.command}')

if __name__ == '__main__':
    main()
