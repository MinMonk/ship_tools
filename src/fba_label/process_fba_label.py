import os
import fitz  # PyMuPDF
import logging
from src.utils.common_utils import root_dir
from src.constants.dir_types import DirType

logger = logging.getLogger(__name__)


def shipment_code_from_filename(filename: str) -> str:
    """从 Amazon 标签文件名中提取货件编号。"""
    stem, _ = os.path.splitext(filename)
    return stem.split("-", 1)[0]


def apply_cover_and_crop(doc, cover_rects, crop_rect):
    """对 PDF 文档应用覆盖和裁剪规则。"""
    white = (1, 1, 1)  # RGB 白色
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        for rect in cover_rects:
            page.draw_rect(rect, color=white, fill=white)
        page.set_cropbox(crop_rect)


def cover_and_crop_pdf(input_file, output_file, cover_rects, crop_rect):
    """覆盖和裁剪 PDF 文件"""
    doc = fitz.open(input_file)
    apply_cover_and_crop(doc, cover_rects, crop_rect)
    doc.save(output_file)
    doc.close()


def save_parity_pdf(src_doc, page_indexes, output_file):
    """按指定页索引保存新 PDF，返回保存的页数。"""
    if not page_indexes:
        return 0

    out_doc = fitz.open()
    for page_index in page_indexes:
        out_doc.insert_pdf(src_doc, from_page=page_index, to_page=page_index)
    out_doc.save(output_file)
    out_doc.close()
    return len(page_indexes)


def process_send_pdf(input_pdf_path, filename, cover_rects, crop_rect):
    """SEND 模式：奇数页为 FBA 标签，偶数页为唛头标签。"""
    fba_output_dir = root_dir(DirType.FBA_Label)
    mark_output_dir = root_dir(DirType.Mark_Label)
    os.makedirs(fba_output_dir, exist_ok=True)
    os.makedirs(mark_output_dir, exist_ok=True)

    shipment_code = shipment_code_from_filename(filename)
    src_doc = fitz.open(input_pdf_path)
    odd_page_indexes = [idx for idx in range(len(src_doc)) if idx % 2 == 0]
    even_page_indexes = [idx for idx in range(len(src_doc)) if idx % 2 == 1]

    fba_doc = fitz.open()
    for page_index in odd_page_indexes:
        fba_doc.insert_pdf(src_doc, from_page=page_index, to_page=page_index)
    apply_cover_and_crop(fba_doc, cover_rects, crop_rect)
    fba_box_num = len(odd_page_indexes)
    fba_output_file = os.path.join(
        fba_output_dir, f"FBA标签_{shipment_code}_{fba_box_num}.pdf"
    )
    fba_doc.save(fba_output_file)
    fba_doc.close()

    mark_box_num = len(even_page_indexes)
    mark_output_file = os.path.join(
        mark_output_dir, f"唛头_{shipment_code}_{mark_box_num}.pdf"
    )
    save_parity_pdf(src_doc, even_page_indexes, mark_output_file)
    src_doc.close()

    logger.debug(f"  已处理 SEND 标签: {filename}")


def process_fba_label(region: str = "US", mode: str = "FIST") -> str:
    input_dir = './shipment_data'
    output_path = root_dir(DirType.FBA_Label)
    os.makedirs(output_path, exist_ok=True)
    mode = mode.upper()

    # 获取区域参数
    logger.info(f"\033[36m▶ 当前模式：{region}\033[0m")
    logger.info(f"\033[36m▶ 标签处理模式：{mode}\033[0m")

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
            if mode == "SEND":
                process_send_pdf(input_pdf_path, filename, cover_rects, crop_rect)
            else:
                output_pdf_path = os.path.join(output_path, filename)
                cover_and_crop_pdf(input_pdf_path, output_pdf_path, cover_rects, crop_rect)
            processed_files += 1
            logger.debug(f"  已处理: {filename}")

    # 完成总结
    logger.info(f"\033[32m✅ 处理完成！共处理 {processed_files} 个文件，输出目录：{output_path} \033[0m")
