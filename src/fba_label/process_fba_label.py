import os
import fitz  # PyMuPDF
import logging
from PIL import Image, ImageChops
from src.utils.common_utils import root_dir
from src.constants.dir_types import DirType

logger = logging.getLogger(__name__)

PT_PER_MM = 72 / 25.4
MARK_LABEL_WIDTH_MM = 100
MARK_LABEL_HEIGHT_MM = 80
MARK_LABEL_MARGIN_MM = 2
MARK_LABEL_PADDING_MM = 1.5
MARK_LABEL_RENDER_SCALE = 3
MARK_LABEL_THRESHOLD = 20


def mm_to_pt(value: float) -> float:
    return value * PT_PER_MM


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


def detect_content_bbox(page, render_scale: float, threshold: int, padding_pt: float):
    """渲染页面后识别非白色内容区域。"""
    pix = page.get_pixmap(matrix=fitz.Matrix(render_scale, render_scale), alpha=False)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    bg = Image.new("RGB", img.size, (255, 255, 255))
    diff = ImageChops.difference(img, bg).convert("L")
    mask = diff.point(lambda x: 255 if x > threshold else 0)
    bbox = mask.getbbox()
    if not bbox:
        return page.rect

    x0, y0, x1, y1 = [v / render_scale for v in bbox]
    crop = fitz.Rect(x0, y0, x1, y1)
    crop.x0 = max(page.rect.x0, crop.x0 - padding_pt)
    crop.y0 = max(page.rect.y0, crop.y0 - padding_pt)
    crop.x1 = min(page.rect.x1, crop.x1 + padding_pt)
    crop.y1 = min(page.rect.y1, crop.y1 + padding_pt)
    return crop


def fit_rect(src_rect, page_width_pt: float, page_height_pt: float, margin_pt: float):
    usable_width = page_width_pt - margin_pt * 2
    usable_height = page_height_pt - margin_pt * 2
    scale = min(usable_width / src_rect.width, usable_height / src_rect.height)
    fitted_width = src_rect.width * scale
    fitted_height = src_rect.height * scale
    x0 = (page_width_pt - fitted_width) / 2
    y0 = (page_height_pt - fitted_height) / 2
    return fitz.Rect(x0, y0, x0 + fitted_width, y0 + fitted_height)


def crop_mark_label_pages(
    src_doc,
    page_indexes,
    output_file,
    width_mm: float = MARK_LABEL_WIDTH_MM,
    height_mm: float = MARK_LABEL_HEIGHT_MM,
    margin_mm: float = MARK_LABEL_MARGIN_MM,
    padding_mm: float = MARK_LABEL_PADDING_MM,
    render_scale: float = MARK_LABEL_RENDER_SCALE,
    threshold: int = MARK_LABEL_THRESHOLD,
):
    """裁剪唛头页空白区域，并适配到 100x80mm 标签纸。"""
    if not page_indexes:
        return 0

    out_doc = fitz.open()
    page_width_pt = mm_to_pt(width_mm)
    page_height_pt = mm_to_pt(height_mm)
    margin_pt = mm_to_pt(margin_mm)
    padding_pt = mm_to_pt(padding_mm)

    for page_index in page_indexes:
        src_page = src_doc[page_index]
        content_bbox = detect_content_bbox(src_page, render_scale, threshold, padding_pt)
        target_page = out_doc.new_page(width=page_width_pt, height=page_height_pt)
        target_rect = fit_rect(content_bbox, page_width_pt, page_height_pt, margin_pt)
        target_page.show_pdf_page(target_rect, src_doc, page_index, clip=content_bbox)

    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    out_doc.save(output_file, garbage=4, deflate=True)
    out_doc.close()
    return len(page_indexes)


def crop_mark_label_pdf(input_path: str, output_path: str, **kwargs):
    """裁剪一个独立唛头 PDF 文件，供测试脚本复用。"""
    src_doc = fitz.open(input_path)
    page_indexes = list(range(len(src_doc)))
    page_count = crop_mark_label_pages(src_doc, page_indexes, output_path, **kwargs)
    src_doc.close()
    return page_count


def save_fba_pages(src_doc, page_indexes, output_file, cover_rects, crop_rect):
    """保存 FBA 标签页，并应用原有裁剪和擦除逻辑。"""
    if not page_indexes:
        return 0

    out_doc = fitz.open()
    for page_index in page_indexes:
        out_doc.insert_pdf(src_doc, from_page=page_index, to_page=page_index)
    apply_cover_and_crop(out_doc, cover_rects, crop_rect)
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    out_doc.save(output_file)
    out_doc.close()
    return len(page_indexes)


def process_send_fba_pdf(input_pdf_path, filename, cover_rects, crop_rect):
    """SEND 模式：奇数页为 FBA 标签。"""
    fba_output_dir = root_dir(DirType.FBA_Label)
    os.makedirs(fba_output_dir, exist_ok=True)

    shipment_code = shipment_code_from_filename(filename)
    src_doc = fitz.open(input_pdf_path)
    odd_page_indexes = [idx for idx in range(len(src_doc)) if idx % 2 == 0]

    fba_box_num = len(odd_page_indexes)
    fba_output_file = os.path.join(
        fba_output_dir, f"FBA标签_{shipment_code}_{fba_box_num}.pdf"
    )
    save_fba_pages(src_doc, odd_page_indexes, fba_output_file, cover_rects, crop_rect)
    src_doc.close()

    logger.debug(f"  已处理 SEND FBA 标签: {filename}")


def process_send_mark_pdf(input_pdf_path, filename):
    """SEND 模式：偶数页为唛头标签，并适配到 100x80mm。"""
    mark_output_dir = root_dir(DirType.Mark_Label)
    os.makedirs(mark_output_dir, exist_ok=True)

    shipment_code = shipment_code_from_filename(filename)
    src_doc = fitz.open(input_pdf_path)
    even_page_indexes = [idx for idx in range(len(src_doc)) if idx % 2 == 1]
    mark_box_num = len(even_page_indexes)
    mark_output_file = os.path.join(
        mark_output_dir, f"唛头_{shipment_code}_{mark_box_num}.pdf"
    )
    crop_mark_label_pages(src_doc, even_page_indexes, mark_output_file)
    src_doc.close()

    logger.debug(f"  已处理 SEND 唛头: {filename}")


def process_fba_label(region: str = "US", mode: str = "FIST", label_type: str = "fba") -> str:
    input_dir = './shipment_data'
    label_type = label_type.lower()
    output_path = root_dir(DirType.Mark_Label if label_type == "mark" else DirType.FBA_Label)
    os.makedirs(output_path, exist_ok=True)
    mode = mode.upper()

    # 获取区域参数
    if label_type == "fba":
        logger.info(f"\033[36m▶ 当前模式：{region}\033[0m")
    logger.info(f"\033[36m▶ 标签处理模式：{mode}\033[0m")

    # 根据区域选择覆盖的矩形区域
    if label_type == "mark":
        cover_rects = []
        crop_rect = None
        logger.info("🔸 操作：裁剪页面")
    else:
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
            if label_type == "mark":
                if mode == "SEND":
                    process_send_mark_pdf(input_pdf_path, filename)
                else:
                    continue
            elif mode == "SEND":
                process_send_fba_pdf(input_pdf_path, filename, cover_rects, crop_rect)
            else:
                output_pdf_path = os.path.join(output_path, filename)
                cover_and_crop_pdf(input_pdf_path, output_pdf_path, cover_rects, crop_rect)
            processed_files += 1
            logger.debug(f"  已处理: {filename}")

    # 完成总结
    logger.info(f"\033[32m✅ 处理完成！共处理 {processed_files} 个文件，输出目录：{output_path} \033[0m")
