import os
import fitz  # PyMuPDF
from datetime import datetime
import sys
import logging
from src.utils.common_utils import root_dir
from src.constants.dir_types import DirType

logger = logging.getLogger(__name__)


# def get_region():
#     """从命令行参数获取区域（US 或 CA），默认为 US"""
#     if len(sys.argv) > 1:
#         region = sys.argv[2].upper()
#         if region not in ["US", "CA"]:
#             print("\033[31m错误：非法参数！仅支持 US 或 CA。\033[0m")
#             sys.exit(1)
#         return region
#     return "US"

def cover_and_crop_pdf(input_file, output_file, cover_rects, crop_rect):
    """覆盖和裁剪 PDF 文件"""
    doc = fitz.open(input_file)
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        white = (1, 1, 1)  # RGB 白色
        for rect in cover_rects:
            page.draw_rect(rect, color=white, fill=white)
        page.set_cropbox(crop_rect)
    doc.save(output_file)
    doc.close()

def process_fba_label(region: str = "US") -> str:
    input_dir = './shipment_data'
    output_path = root_dir(DirType.FBA_Label)
    os.makedirs(output_path, exist_ok=True)

    # 获取区域参数
    logger.info(f"\033[36m▶ 当前模式：{region}\033[0m")

    # 根据区域选择覆盖的矩形区域
    if region == "US":
        cover_rects = [fitz.Rect(27, 33, 130, 42)]
        logger.info("🔸 操作：裁剪页面 + 擦除目的地公司名称")
    else:
        cover_rects = [fitz.Rect(27, 33, 130, 42), fitz.Rect(150, 24, 280, 70)]
        logger.info("🔸 操作：裁剪页面 + 擦除目的地公司名称及发货地信息")

    # 固定裁剪区域
    crop_rect = fitz.Rect(0, 0, 306, 230)

    # 处理文件
    processed_files = 0
    for filename in os.listdir(input_dir):
        if filename.endswith('.pdf'):
            input_pdf_path = os.path.join(input_dir, filename)
            output_pdf_path = os.path.join(output_path, filename)
            cover_and_crop_pdf(input_pdf_path, output_pdf_path, cover_rects, crop_rect)
            processed_files += 1
            logger.debug(f"  已处理: {filename}")

    # 完成总结
    logger.info(f"\033[32m✅ 处理完成！共处理 {processed_files} 个文件，输出目录：{output_path} \033[0m")