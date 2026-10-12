"""本地智能填表客户端的资料库服务。

开发启动：uvicorn main:app --reload --port 8999
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import sqlite3
import sys
import tempfile
import time
import urllib.error
import urllib.request
import uuid
from copy import copy, deepcopy
from contextlib import contextmanager
from datetime import UTC, datetime
from pathlib import Path

from docx import Document
from fastapi import FastAPI, HTTPException, status
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
from pypdf import PdfReader
from pydantic import BaseModel, Field
import uvicorn

APP_FOLDER_NAME = "fill_form"
KNOWLEDGE_EXTENSIONS = {".docx", ".xlsx", ".pdf"}
FORM_EXTENSIONS = {".docx", ".xlsx"}
TEXT_LIMIT = 100_000
LLM_BATCH_TIMEOUT_SECONDS = 180
FORM_DETECTOR_VERSION = 6
DEFAULT_WRITING_STYLE = "专业、简洁，突出事实与成果；不得编造资料库中不存在的事实。"
DEFAULT_STYLE_ID = "default"


def default_data_root() -> Path:
    """与 Electron 默认 userData 目录保持一致，可由环境变量覆盖。"""
    if custom_root := os.getenv("FILL_FORM_DATA_DIR"):
        return Path(custom_root).expanduser().resolve()
    if sys.platform == "win32":
        return Path(os.getenv("APPDATA", Path.home())) / APP_FOLDER_NAME / "data"
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / APP_FOLDER_NAME / "data"
    return Path(os.getenv("XDG_DATA_HOME", Path.home() / ".local" / "share")) / APP_FOLDER_NAME / "data"


DATA_ROOT = default_data_root()
KNOWLEDGE_DIR = DATA_ROOT / "knowledge"
FORMS_DIR = DATA_ROOT / "forms"
OUTPUTS_DIR = DATA_ROOT / "outputs"
# Paddle 的 C++ 推理引擎在部分 Windows 环境无法可靠读取含中文的模型路径。
# 安装版的 Electron userData 可能以产品中文名命名，因此 OCR 缓存固定使用
# fill_form 这个 ASCII 目录；资料、表单等其余本地数据仍在 DATA_ROOT 中。
OCR_DIR = (
    Path(os.getenv("APPDATA", Path.home())) / APP_FOLDER_NAME / "data" / "ocr"
    if sys.platform == "win32"
    else DATA_ROOT / "ocr"
)
DATABASE_PATH = DATA_ROOT / "app.db"
SETTINGS_PATH = DATA_ROOT / "settings.json"
# OCR 是耗时任务。进度仅用于当前运行中的本地客户端展示，无需写入资料库。
OCR_PROGRESS: dict[str, dict] = {}


def set_ocr_progress(file_id: str, state: str, progress: int, message: str) -> None:
    OCR_PROGRESS[file_id] = {
        "state": state,
        "progress": max(0, min(progress, 100)),
        "message": message,
    }


def ensure_storage() -> None:
    for directory in (DATA_ROOT, KNOWLEDGE_DIR, FORMS_DIR, OUTPUTS_DIR, OCR_DIR):
        directory.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(DATABASE_PATH) as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS knowledge_files (
                id TEXT PRIMARY KEY,
                original_name TEXT NOT NULL,
                extension TEXT NOT NULL,
                source_path TEXT NOT NULL,
                stored_path TEXT NOT NULL,
                sha256 TEXT NOT NULL UNIQUE,
                imported_at TEXT NOT NULL,
                parse_status TEXT NOT NULL,
                parse_error TEXT,
                extracted_text TEXT NOT NULL DEFAULT ''
            )
            """
        )
        connection.execute(
            "CREATE INDEX IF NOT EXISTS idx_knowledge_files_source_path ON knowledge_files(source_path)"
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS forms (
                id TEXT PRIMARY KEY,
                original_name TEXT NOT NULL,
                extension TEXT NOT NULL,
                source_path TEXT NOT NULL,
                stored_path TEXT NOT NULL,
                sha256 TEXT NOT NULL UNIQUE,
                imported_at TEXT NOT NULL,
                detector_version INTEGER NOT NULL DEFAULT 1
            )
            """
        )
        form_columns = {row[1] for row in connection.execute("PRAGMA table_info(forms)").fetchall()}
        if "detector_version" not in form_columns:
            connection.execute("ALTER TABLE forms ADD COLUMN detector_version INTEGER NOT NULL DEFAULT 1")
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS form_fields (
                id TEXT PRIMARY KEY,
                form_id TEXT NOT NULL,
                raw_label TEXT NOT NULL,
                field_type TEXT NOT NULL,
                confidence REAL NOT NULL,
                evidence TEXT NOT NULL,
                location_json TEXT NOT NULL,
                current_value TEXT NOT NULL DEFAULT '',
                needs_ai_generation INTEGER NOT NULL DEFAULT 0,
                FOREIGN KEY(form_id) REFERENCES forms(id) ON DELETE CASCADE
            )
            """
        )
        existing_columns = {
            row[1] for row in connection.execute("PRAGMA table_info(form_fields)").fetchall()
        }
        for column, definition in (
            ("suggested_value", "TEXT NOT NULL DEFAULT ''"),
            ("suggestion_reason", "TEXT NOT NULL DEFAULT ''"),
            ("source_ids_json", "TEXT NOT NULL DEFAULT '[]'"),
                ("generation_status", "TEXT NOT NULL DEFAULT 'not_generated'"),
            ("final_value", "TEXT NOT NULL DEFAULT ''"),
            ("instruction", "TEXT NOT NULL DEFAULT ''"),
        ):
            if column not in existing_columns:
                connection.execute(f"ALTER TABLE form_fields ADD COLUMN {column} {definition}")
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS exports (
                id TEXT PRIMARY KEY,
                form_id TEXT NOT NULL,
                form_name TEXT NOT NULL,
                output_path TEXT NOT NULL,
                output_name TEXT NOT NULL,
                filled_field_count INTEGER NOT NULL,
                missing_field_count INTEGER NOT NULL,
                created_at TEXT NOT NULL
            )
            """
        )


def load_settings() -> dict:
    if not SETTINGS_PATH.exists():
        return {}
    try:
        return json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise HTTPException(status_code=500, detail="本地模型配置文件无法读取") from error


def save_settings(settings: dict) -> None:
    ensure_storage()
    temporary_path = SETTINGS_PATH.with_suffix(".tmp")
    temporary_path.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary_path.replace(SETTINGS_PATH)


def current_outputs_dir() -> Path:
    """导出目录：优先使用客户端配置的位置，没配置就用默认目录；目录不存在时自动创建。"""
    configured = str(load_settings().get("output_dir") or "").strip()
    directory = Path(configured).expanduser().resolve() if configured else OUTPUTS_DIR
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def load_writing_styles() -> dict:
    """写作风格列表 + 当前选中的风格。

    老版本只存了一条 writing_style 字符串，这里会自动迁移成列表里的第一条。
    """
    settings = load_settings()
    raw_styles = settings.get("writing_styles")
    if not isinstance(raw_styles, list) or not raw_styles:
        legacy = str(settings.get("writing_style") or "").strip() or DEFAULT_WRITING_STYLE
        raw_styles = [{"id": DEFAULT_STYLE_ID, "name": "默认写作风格", "content": legacy}]

    styles = [
        {
            "id": str(item.get("id") or "").strip() or DEFAULT_STYLE_ID,
            "name": str(item.get("name") or "").strip() or "未命名风格",
            "content": str(item.get("content") or "").strip() or DEFAULT_WRITING_STYLE,
        }
        for item in raw_styles
        if isinstance(item, dict)
    ]
    if not styles:
        styles = [{"id": DEFAULT_STYLE_ID, "name": "默认写作风格", "content": DEFAULT_WRITING_STYLE}]

    active_id = str(settings.get("active_style_id") or "").strip()
    if not any(item["id"] == active_id for item in styles):
        active_id = styles[0]["id"]
    return {"styles": styles, "active_id": active_id}


def current_writing_style() -> str:
    """当前选中风格的内容；找不到时回退到默认风格。"""
    data = load_writing_styles()
    for item in data["styles"]:
        if item["id"] == data["active_id"]:
            return item["content"]
    return DEFAULT_WRITING_STYLE


@contextmanager
def database_connection():
    ensure_storage()
    connection = sqlite3.connect(DATABASE_PATH)
    connection.row_factory = sqlite3.Row
    try:
        yield connection
        connection.commit()
    finally:
        connection.close()


def sha256_for_file(file_path: Path) -> str:
    digest = hashlib.sha256()
    with file_path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def limit_text(parts: list[str]) -> str:
    return "\n".join(part.strip() for part in parts if part and part.strip())[:TEXT_LIMIT]


def extract_docx_text(file_path: Path) -> str:
    document = Document(file_path)
    content = [paragraph.text for paragraph in document.paragraphs]
    for table in document.tables:
        for row in table.rows:
            content.append(" | ".join(cell.text for cell in row.cells))
    return limit_text(content)


def extract_xlsx_text(file_path: Path) -> str:
    workbook = load_workbook(file_path, read_only=True, data_only=True)
    content: list[str] = []
    try:
        for worksheet in workbook.worksheets:
            content.append(f"工作表：{worksheet.title}")
            for row in worksheet.iter_rows(values_only=True):
                values = [str(value) for value in row if value is not None and str(value).strip()]
                if values:
                    content.append(" | ".join(values))
                if sum(map(len, content)) >= TEXT_LIMIT:
                    return limit_text(content)
    finally:
        workbook.close()
    return limit_text(content)


def extract_pdf_text(file_path: Path) -> str:
    """提取可复制文字的 PDF 内容；扫描件没有文本层时提示用户使用 OCR。"""
    reader = PdfReader(str(file_path))
    if reader.is_encrypted and reader.decrypt("") == 0:
        raise ValueError("PDF 已加密，当前不支持需要密码的资料文件")

    content: list[str] = []
    for page_number, page in enumerate(reader.pages, start=1):
        text = (page.extract_text() or "").strip()
        if text:
            content.append(f"第 {page_number} 页\n{text}")
        if sum(map(len, content)) >= TEXT_LIMIT:
            break
    extracted_text = limit_text(content)
    if not extracted_text:
        raise ValueError("未提取到可复制文字；该 PDF 可能是扫描件，请点击“OCR 识别”")
    return extracted_text


def _ocr_dependency_error(package_name: str) -> HTTPException:
    return HTTPException(
        status_code=503,
        detail=(
            f"扫描 PDF 识别组件尚未安装（缺少 {package_name}）。"
            "请在 PyCharm 当前项目解释器中执行 requirements.txt 的安装后重试。"
        ),
    )


def clear_incomplete_ocr_models(error: Exception) -> bool:
    """仅清理 PaddleOCR 已下载但关键文件缺失的模型目录，供下一次自动重下。"""
    message = str(error)
    if "Cannot open file" not in message or "official_models" not in message:
        return False
    model_root = (OCR_DIR / "paddlex-cache" / "official_models").resolve()
    # PaddleX 报错路径通常类似 official_models\PP-OCRv6_medium_det\inference.json。
    model_names = set(re.findall(r"official_models[\\/](PP-OCR[^\\/]+)", message))
    removed = False
    for model_name in model_names:
        candidate = (model_root / model_name).resolve()
        if candidate.parent == model_root and candidate.is_dir():
            shutil.rmtree(candidate)
            removed = True
    return removed


def create_paddle_ocr(progress_callback=None):
    """按需加载 OCR，首次创建时由 PaddleOCR 从官方源下载所需模型。

    不在模块加载时导入，保证未安装 OCR 依赖的用户仍可使用 Word、Excel
    和文本型 PDF。模型缓存指定到应用数据目录，避免散落在用户目录下。
    """
    os.environ.setdefault("PADDLE_PDX_CACHE_HOME", str(OCR_DIR / "paddlex-cache"))
    os.environ.setdefault("PADDLEOCR_HOME", str(OCR_DIR / "paddleocr-cache"))
    # PaddlePaddle 3.3.x 的 Windows CPU 版存在 PIR 与 oneDNN 的兼容问题；
    # 必须在导入 Paddle 前关闭 PIR，避免原生层直接中断识别进程。
    os.environ.setdefault("FLAGS_enable_pir_api", "0")
    try:
        from paddleocr import PaddleOCR
    except ImportError as error:
        raise _ocr_dependency_error("paddleocr / paddlepaddle") from error

    def build_ocr():
        return PaddleOCR(
            use_doc_orientation_classify=False,
            use_doc_unwarping=False,
            use_textline_orientation=False,
            engine="paddle",
            # Paddle 3.3.x 在部分 Windows CPU 上启用 oneDNN 会触发
            # ConvertPirAttribute2RuntimeAttribute 错误。关闭该加速更稳定。
            enable_mkldnn=False,
        )

    try:
        return build_ocr()
    except Exception as error:
        # 网络中断时 PaddleX 有时会留下同名目录，却没有完整下载 inference.json。
        # 此时“文件已存在”会让后续重试永远失败，因此只清理报错点对应的模型并重试一次。
        if clear_incomplete_ocr_models(error):
            if progress_callback:
                progress_callback(14, "检测到 OCR 模型下载不完整，正在清理并重新下载…")
            try:
                return build_ocr()
            except Exception as retry_error:
                error = retry_error
        raise HTTPException(
            status_code=502,
            detail=(
                "无法从 PaddleOCR 官方源准备 OCR 模型。请检查网络后重试；"
                f"具体原因：{error}"
            ),
        ) from error


def ocr_result_text(result: object) -> list[str]:
    """兼容 PaddleOCR 3.x 的结果对象，读取文字和对应的置信度。"""
    data = result
    json_method = getattr(result, "json", None)
    if callable(json_method):
        try:
            data = json.loads(json_method())
        except (TypeError, ValueError):
            data = result
    if not isinstance(data, dict):
        data = getattr(result, "res", {})
    if not isinstance(data, dict):
        return []

    texts = data.get("rec_texts", [])
    scores = data.get("rec_scores", [])
    if not isinstance(texts, list):
        return []
    return [
        str(text).strip()
        for index, text in enumerate(texts)
        if str(text).strip()
        and (index >= len(scores) or not isinstance(scores[index], (int, float)) or scores[index] >= 0.5)
    ]


def extract_scanned_pdf_text(file_path: Path, progress_callback=None) -> str:
    """将扫描版 PDF 渲染成页面图片，再使用 PaddleOCR 识别文字。"""
    try:
        import fitz
    except ImportError as error:
        raise _ocr_dependency_error("PyMuPDF") from error

    if progress_callback:
        progress_callback(12, "正在准备 OCR 引擎；首次使用时正在从官方源下载模型…")
    ocr = create_paddle_ocr(progress_callback)
    try:
        document = fitz.open(str(file_path))
    except Exception as error:
        raise ValueError(f"无法打开 PDF：{error}") from error

    content: list[str] = []
    try:
        with tempfile.TemporaryDirectory(prefix="ocr-", dir=OCR_DIR) as temporary_dir:
            total_pages = len(document)
            if progress_callback:
                progress_callback(28, f"模型准备完成，开始识别，共 {total_pages} 页…")
            for page_number, page in enumerate(document, start=1):
                # 约 144 DPI：兼顾中文识别率与首次使用时的处理速度。
                pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                image_path = Path(temporary_dir) / f"page-{page_number}.png"
                pixmap.save(str(image_path))

                page_text: list[str] = []
                for result in ocr.predict(str(image_path)):
                    page_text.extend(ocr_result_text(result))
                if page_text:
                    content.append(f"第 {page_number} 页\n" + "\n".join(page_text))
                if progress_callback:
                    progress_callback(
                        28 + int(page_number / max(total_pages, 1) * 67),
                        f"正在识别第 {page_number}/{total_pages} 页…",
                    )
                if sum(map(len, content)) >= TEXT_LIMIT:
                    break
    finally:
        document.close()

    extracted_text = limit_text(content)
    if not extracted_text:
        raise ValueError("OCR 未识别出文字，请确认 PDF 清晰、方向正确且不是纯空白页面")
    return extracted_text


def extract_text(file_path: Path) -> str:
    if file_path.suffix.lower() == ".docx":
        return extract_docx_text(file_path)
    if file_path.suffix.lower() == ".xlsx":
        return extract_xlsx_text(file_path)
    if file_path.suffix.lower() == ".pdf":
        return extract_pdf_text(file_path)
    raise ValueError("仅支持 .docx、.xlsx 与 .pdf 文件")


def infer_field_type(label: str) -> tuple[str, bool]:
    """仅根据标题给出展示类型；未知字段仍会保留原始标题供 LLM 理解。"""
    normalized = re.sub(r"\s+", "", label).lower()
    if any(word in normalized for word in ("感言", "心得", "评价", "总结", "计划", "描述", "说明", "原因", "经历", "贡献", "事迹", "内容")):
        return "long_text", True
    if any(word in normalized for word in ("电话", "手机", "联系方式")):
        return "phone", False
    if "邮箱" in normalized or "email" in normalized:
        return "email", False
    if any(word in normalized for word in ("日期", "时间", "出生年月")):
        return "date", False
    return "short_text", False


# 表头说明、示例行等明显不是待填字段的标题
IGNORED_LABELS = {
    "例", "示例", "举例", "样例", "示例填写", "填写示例",
    "填写说明", "注意事项", "备注说明", "说明文字", "提示",
}


def is_valid_label(label: str) -> bool:
    """过滤表头与示例行，避免把“例”这类文字当成待填字段。"""
    normalized = label.strip().strip("：:").strip()
    if not normalized or normalized in IGNORED_LABELS:
        return False
    # 标题至少要有一个汉字或英文字母，纯数字/符号不算
    if not re.search(r"[\u4e00-\u9fffA-Za-z]", normalized):
        return False
    # 单个汉字通常是表头残留
    return len(normalized) >= 2


def cell_label(worksheet, row: int, column: int) -> tuple[str, str]:
    """取得单元格的标题文字，并返回来源坐标；合并单元格只认左上角锚点。"""
    for merged_range in worksheet.merged_cells.ranges:
        if merged_range.min_row <= row <= merged_range.max_row and merged_range.min_col <= column <= merged_range.max_col:
            value = worksheet.cell(merged_range.min_row, merged_range.min_col).value
            text = "" if value is None else str(value).strip()
            return text, f"{merged_range.min_row}:{merged_range.min_col}"
    value = worksheet.cell(row, column).value
    text = "" if value is None else str(value).strip()
    return text, f"{row}:{column}"


def is_non_anchor_merged_cell(worksheet, row: int, column: int) -> bool:
    for merged_range in worksheet.merged_cells.ranges:
        if merged_range.min_row <= row <= merged_range.max_row and merged_range.min_col <= column <= merged_range.max_col:
            return (row, column) != (merged_range.min_row, merged_range.min_col)
    return False


def has_input_cell_cue(cell) -> bool:
    """仅将有边框或底色的空白单元格视为表单输入区域。

    普通工作表经常在表格右侧或下方保留大量空白格；这些格不应因邻近文字而被自动填写。
    """
    fill = cell.fill
    has_fill = bool(fill and fill.fill_type and fill.fill_type != "none")
    has_border = any(
        side.style for side in (cell.border.left, cell.border.right, cell.border.top, cell.border.bottom)
    )
    return has_fill or has_border


def find_example_rows(worksheet, max_row: int, max_column: int) -> set[int]:
    """找出以“例/示例”等标记开头的示例行；这类行只做示范，不作为待填行。"""
    example_rows: set[int] = set()
    for row in range(1, max_row + 1):
        for column in range(1, max_column + 1):
            value = worksheet.cell(row, column).value
            if value is None or not str(value).strip():
                continue
            if str(value).strip().strip("：:") in IGNORED_LABELS:
                example_rows.add(row)
            break
    return example_rows


def row_filled_count(worksheet, row: int, max_column: int) -> int:
    """统计一行里非空单元格的数量。"""
    return sum(
        1 for column in range(1, max_column + 1)
        if str(worksheet.cell(row, column).value or "").strip()
    )


def find_title_rows(worksheet, max_row: int, max_column: int) -> set[int]:
    """识别表格最上方的标题行：整行只写了 1 个单元格，且下一行是表头（多个非空单元格）。"""
    for row in range(1, max_row):
        filled = row_filled_count(worksheet, row, max_column)
        if filled == 0:
            continue
        # 只有“孤立的一格 + 下面紧跟表头”才算标题，避免误伤正常的键值对首行
        if filled == 1 and row_filled_count(worksheet, row + 1, max_column) >= 2:
            return {row}
    return set()


TABLE_CONTROL_LABELS = {"序号", "编号", "序列号", "no", "no."}


def find_tabular_header(worksheet, max_row: int, max_column: int, example_rows: set[int]) -> tuple[int, list[int]] | None:
    """识别“表头 + 多条记录”的明细表。

    明细表的每一列标题应来自同一行，不能把上一条记录里的“小红”“湖南”
    之类的数据误当成下一行的字段名。至少四列可有效区分常见的键值对表单。
    """
    for row in range(1, max_row):
        header_columns = []
        header_sources: set[str] = set()
        for column in range(1, max_column + 1):
            text, source = cell_label(worksheet, row, column)
            # 合并标题单元格会在这一行的每一列都返回同一个锚点；它只能算一个标题，
            # 不能误判为多列表头。
            if source in header_sources:
                continue
            if is_valid_label(text):
                header_columns.append(column)
                header_sources.add(source)
        if len(header_columns) < 4:
            continue

        # 标题行下方应至少有一行填过多列数据，才按“明细表”处理。
        required_values = max(2, len(header_columns) // 2)
        for data_row in range(row + 1, min(max_row, row + 5) + 1):
            if data_row in example_rows:
                continue
            value_count = sum(
                1 for column in header_columns
                if str(worksheet.cell(data_row, column).value or "").strip()
            )
            if value_count >= required_values:
                return row, header_columns
    return None


def find_tabular_target_row(
    worksheet, header_row: int, header_columns: list[int], max_row: int, example_rows: set[int]
) -> tuple[int, list[int]] | None:
    """选择第一条未完成记录：优先部分已填行，其次完全空白行。"""
    empty_row_candidate: tuple[int, list[int]] | None = None
    for row in range(header_row + 1, max_row + 1):
        if row in example_rows:
            continue
        blanks = [
            column for column in header_columns
            if worksheet.cell(row, column).value is None
            and not is_non_anchor_merged_cell(worksheet, row, column)
            and has_input_cell_cue(worksheet.cell(row, column))
        ]
        if not blanks:
            continue
        filled_count = sum(
            1 for column in header_columns
            if str(worksheet.cell(row, column).value or "").strip()
        )
        if filled_count:
            return row, blanks
        if empty_row_candidate is None:
            empty_row_candidate = (row, blanks)
    return empty_row_candidate


def detect_xlsx_fields(file_path: Path) -> list[dict]:
    workbook = load_workbook(file_path, data_only=False)
    fields: list[dict] = []
    seen_label_sources: set[tuple[str, str]] = set()
    try:
        for worksheet in workbook.worksheets:
            # 限制扫描范围，避免用户的格式化空白列被误当作表单字段。
            max_row, max_column = min(worksheet.max_row, 500), min(worksheet.max_column, 100)
            # 示例行（行首写着“例”“示例”）和表格顶部的标题行都不参与填写
            example_rows = find_example_rows(worksheet, max_row, max_column)
            skip_rows = example_rows | find_title_rows(worksheet, max_row, max_column)
            # 项目经历表是可重复区域：一个项目对应一整行，优先于普通明细表规则。
            tabular_cells: set[tuple[int, int]] = set()
            if project_table := find_xlsx_project_table(worksheet, max_row, max_column):
                fields.append({
                    "label": "项目经历",
                    "field_type": "repeat_table",
                    "confidence": 0.95,
                    "evidence": (
                        f"识别到第 {project_table['header_row']} 行项目经历表头；"
                        "每个项目会单独填写一行，并可自动扩展预留行"
                    ),
                    "instruction": "请从资料中提取项目经历；每个项目返回一行，不得编造。",
                    "location": {"file_type": "xlsx", "sheet": worksheet.title, **project_table},
                    "needs_ai_generation": True,
                })
                for row in project_table["template_row_indexes"]:
                    for column in range(1, max_column + 1):
                        tabular_cells.add((row, column))
            # 对普通明细表仅填写第一条未完成记录；整块记录区不再进入通用的“左/上邻格”规则。
            elif tabular_header := find_tabular_header(worksheet, max_row, max_column, example_rows):
                header_row, header_columns = tabular_header
                target = find_tabular_target_row(
                    worksheet, header_row, header_columns, max_row, example_rows
                )
                for row in range(header_row + 1, max_row + 1):
                    for column in header_columns:
                        tabular_cells.add((row, column))
                if target:
                    target_row, blank_columns = target
                    for column in blank_columns:
                        label, label_source = cell_label(worksheet, header_row, column)
                        if not is_valid_label(label) or label.strip().lower() in TABLE_CONTROL_LABELS:
                            continue
                        source_key = (worksheet.title, f"table:{header_row}:{label_source}")
                        if source_key in seen_label_sources:
                            continue
                        seen_label_sources.add(source_key)
                        field_type, needs_ai_generation = infer_field_type(label)
                        fields.append({
                            "label": label.rstrip("：:"),
                            "field_type": field_type,
                            "confidence": 0.96,
                            "evidence": (
                                f"明细表第 {header_row} 行为表头；"
                                f"第 {target_row} 行是第一条待补充记录"
                            ),
                            "location": {
                                "file_type": "xlsx", "sheet": worksheet.title,
                                "cell": worksheet.cell(target_row, column).coordinate,
                            },
                            "needs_ai_generation": needs_ai_generation,
                        })
            for row in range(1, max_row + 1):
                if row in skip_rows:
                    continue
                for column in range(1, max_column + 1):
                    cell = worksheet.cell(row, column)
                    if (
                        cell.value is not None
                        or is_non_anchor_merged_cell(worksheet, row, column)
                        or not has_input_cell_cue(cell)
                        or (row, column) in tabular_cells
                    ):
                        continue
                    # 只认紧邻的标题：左边一格或上边一格，避免整片空白网格都变成待填项。
                    candidates = []
                    if column > 1:
                        text, source = cell_label(worksheet, row, column - 1)
                        if text:
                            candidates.append((text, source, 0.94, "同一行左侧相邻单元格为标题"))
                    # 向上找表头时跳过示例行，让真实数据行仍能对上表头
                    above_row = row - 1
                    while above_row in example_rows and above_row > 1:
                        above_row -= 1
                    if above_row > 0:
                        text, source = cell_label(worksheet, above_row, column)
                        if text:
                            candidates.append((text, source, 0.86, "同一列上方相邻单元格为标题"))
                    # 左侧是序号、示例文字时，自动改用上方的表头
                    picked = next((item for item in candidates if is_valid_label(item[0])), None)
                    if picked is None:
                        continue
                    label, label_source, confidence, evidence = picked
                    # 合并标题会让多个空白格共用同一个标题，这里只保留最先遇到的一个。
                    source_key = (worksheet.title, label_source)
                    if source_key in seen_label_sources:
                        continue
                    seen_label_sources.add(source_key)
                    field_type, needs_ai_generation = infer_field_type(label)
                    fields.append({
                        "label": label,
                        "field_type": field_type,
                        "confidence": confidence,
                        "evidence": evidence,
                        "location": {"file_type": "xlsx", "sheet": worksheet.title, "cell": cell.coordinate},
                        "needs_ai_generation": needs_ai_generation,
                    })
    finally:
        workbook.close()
    return fields


def docx_row_filled_count(row) -> int:
    """统计 Word 表格一行里非空单元格的数量；合并单元格只算一次。"""
    seen: set[int] = set()
    count = 0
    for cell in row.cells:
        identity = id(cell._tc)
        if identity in seen:
            continue
        seen.add(identity)
        if cell.text.strip():
            count += 1
    return count


def find_docx_title_rows(table) -> set[int]:
    """识别 Word 表格最上方的标题行：整行只写了 1 个单元格，且下一行是表头。"""
    rows = table.rows
    for row_index in range(len(rows) - 1):
        filled = docx_row_filled_count(rows[row_index])
        if filled == 0:
            continue
        # 只有“孤立的一格 + 下面紧跟表头”才算标题，避免误伤正常的键值对首行
        if filled == 1 and docx_row_filled_count(rows[row_index + 1]) >= 2:
            return {row_index}
        return set()
    return set()


CHECKBOX_MARKS = "□☐☑☒■√"
CHECKBOX_OPTION_PATTERN = re.compile(
    rf"[{re.escape(CHECKBOX_MARKS)}]\s*(.+?)(?=(?:\s*[{re.escape(CHECKBOX_MARKS)}])|$)"
)


def extract_checkbox_options(text: str) -> list[str]:
    """提取 Word 里“□ 是  □ 否”形式的单选项。"""
    options: list[str] = []
    for match in CHECKBOX_OPTION_PATTERN.finditer(text.strip()):
        option = re.sub(r"\s+", " ", match.group(1)).strip(" \t\n\r、,，;；")
        if option and len(option) <= 40 and option not in options:
            options.append(option)
    return options if len(options) >= 2 else []


def normalize_checkbox_choice(value: str, options: list[str]) -> str:
    """将模型或用户输入规范为候选项之一；无匹配时返回空字符串。"""
    normalized = re.sub(r"\s+", "", value or "").strip("'\"“”‘’。；;，,")
    for option in options:
        if normalized == re.sub(r"\s+", "", option):
            return option
    # 模型偶尔会返回“视情况而定（项目需要时）”，此时取唯一命中的候选项。
    matches = [option for option in options if re.sub(r"\s+", "", option) in normalized]
    return matches[0] if len(matches) == 1 else ""


def mark_docx_checkbox(cell, selected_option: str, options: list[str]) -> None:
    """保留 Word 单元格文本，仅更新每个候选项前的勾选符号。"""
    selected_option = normalize_checkbox_choice(selected_option, options)
    if not selected_option:
        return
    for paragraph in cell.paragraphs:
        text = paragraph.text
        for option in sorted(options, key=len, reverse=True):
            pattern = re.compile(
                rf"[{re.escape(CHECKBOX_MARKS)}](?P<space>\s*){re.escape(option)}"
                rf"(?=(?:\s*[{re.escape(CHECKBOX_MARKS)}])|\s*$)"
            )
            mark = "☑" if option == selected_option else "□"
            text = pattern.sub(lambda match: f"{mark}{match.group('space')}{option}", text)
        paragraph.text = text


# 项目经历表属于“可重复区域”：一条资料对应表格中的一整行，不能按普通单元格
# 逐个填，否则多个项目会被拼进第一个空白格。仅识别常见的项目经历表头，避免把
# 其他普通表格误判为可扩展表格。
PROJECT_COLUMN_ALIASES = {
    "project_name": ("项目名称", "项目名", "项目经历", "项目"),
    "project_period": ("项目时间", "项目周期", "起止时间", "时间"),
    "project_role": ("担任角色", "项目角色", "职责", "角色"),
    "project_result": ("核心贡献或成果", "核心贡献", "项目成果", "主要成果", "工作内容", "项目简介"),
}


def project_column_key(label: str) -> str | None:
    normalized = re.sub(r"\s+", "", label).strip("：:")
    for key, aliases in PROJECT_COLUMN_ALIASES.items():
        if normalized in aliases:
            return key
    return None


def find_xlsx_project_table(worksheet, max_row: int, max_column: int) -> dict | None:
    """识别 Excel 中带“项目名称”等列的项目经历明细表。"""
    for header_row in range(1, max_row):
        columns: list[dict] = []
        serial_column: int | None = None
        for column in range(1, max_column + 1):
            label, _ = cell_label(worksheet, header_row, column)
            if label.strip().lower() in TABLE_CONTROL_LABELS:
                serial_column = column
            if key := project_column_key(label):
                columns.append({"key": key, "label": label.strip(), "column": column})
        keys = {column["key"] for column in columns}
        if len(keys) < 3 or "project_name" not in keys:
            continue

        data_start = header_row + 1
        template_rows: list[int] = []
        relevant_columns = [column["column"] for column in columns]
        for row in range(data_start, max_row + 1):
            first_text = str(worksheet.cell(row, 1).value or "").strip()
            # “说明/备注”行不是数据模板，插入新行必须放在它的上面。
            if first_text in {"说明", "备注", "注", "注："}:
                break
            has_table_format = any(
                has_input_cell_cue(worksheet.cell(row, column)) for column in relevant_columns
            )
            if not has_table_format:
                break
            template_rows.append(row)
        if template_rows:
            return {
                "header_row": header_row,
                "data_start_row": data_start,
                "template_row_indexes": template_rows,
                "columns": columns,
                "serial_column": serial_column,
            }
    return None


def copy_xlsx_row_style(worksheet, source_row: int, target_row: int, max_column: int) -> None:
    """复制模板行的视觉样式；Excel 插入行本身不会继承样式。"""
    worksheet.row_dimensions[target_row].height = worksheet.row_dimensions[source_row].height
    for column in range(1, max_column + 1):
        source = worksheet.cell(source_row, column)
        target = worksheet.cell(target_row, column)
        if source.has_style:
            target._style = copy(source._style)
        if source.number_format:
            target.number_format = source.number_format
        if source.alignment:
            target.alignment = copy(source.alignment)
        if source.protection:
            target.protection = copy(source.protection)


def find_docx_project_table(table) -> dict | None:
    """找到 Word 项目经历表的表头和可复制的数据行。"""
    for header_row_index, row in enumerate(table.rows):
        columns: list[dict] = []
        seen_cells: set[int] = set()
        for cell_index, cell in enumerate(row.cells):
            identity = id(cell._tc)
            if identity in seen_cells:
                continue
            seen_cells.add(identity)
            key = project_column_key(cell.text)
            if key:
                columns.append({"key": key, "label": cell.text.strip(), "cell_index": cell_index})
        keys = {column["key"] for column in columns}
        # 至少三列且包含项目名称，才能确信这是项目经历明细表。
        if len(keys) < 3 or "project_name" not in keys:
            continue
        data_start = header_row_index + 1
        if data_start >= len(table.rows):
            continue
        # 表头之后连续的行都视为模板/预留数据行；通常只有 1～2 行。
        # 仅保留列数相同的行，防止把表格尾部的说明行一起复制。
        template_rows: list[int] = []
        expected_cells = len(row.cells)
        for row_index in range(data_start, len(table.rows)):
            candidate = table.rows[row_index]
            if len(candidate.cells) != expected_cells:
                break
            template_rows.append(row_index)
        if not template_rows:
            continue
        return {
            "header_row_index": header_row_index,
            "data_start_row": data_start,
            "template_row_indexes": template_rows,
            "columns": columns,
        }
    return None


def normalize_repeat_rows(value: object, columns: list[dict]) -> list[dict]:
    """验证模型/用户提交的项目行，只保留模板中存在的列。"""
    raw_rows = value
    if isinstance(value, str):
        try:
            raw_rows = json.loads(value)
        except json.JSONDecodeError:
            return []
    if not isinstance(raw_rows, list):
        return []
    allowed = {column["key"] for column in columns}
    normalized_rows: list[dict] = []
    for row in raw_rows:
        if not isinstance(row, dict):
            continue
        normalized: dict[str, str] = {}
        for column in columns:
            key = column["key"]
            # 允许 AI 使用列 key 或中文表头作为对象键，便于兼容不同模型。
            cell_value = row.get(key, row.get(column["label"], ""))
            normalized[key] = str(cell_value or "").strip()
        if any(normalized[key] for key in allowed):
            normalized_rows.append(normalized)
    return normalized_rows


def detect_docx_fields(file_path: Path) -> list[dict]:
    document = Document(file_path)
    fields: list[dict] = []
    # 用对象引用（而不是 id）记录已处理的单元格：lxml 的代理对象会被回收，
    # id 可能被复用，导致同一次识别里莫名漏掉字段。
    seen_cells: list = []
    seen_labels: list = []
    for table_index, table in enumerate(document.tables):
        if repeat_table := find_docx_project_table(table):
            fields.append({
                "label": "项目经历",
                "field_type": "repeat_table",
                "confidence": 0.94,
                "evidence": "识别到“项目名称”等项目经历表头，将按每个项目单独填充一行",
                "instruction": "请从资料中提取项目经历；每个项目返回一行，不得编造。",
                "location": {"file_type": "docx", "table_index": table_index, **repeat_table},
                "needs_ai_generation": True,
            })
            # 该表格的空白格由重复区域统一处理，不能再识别成普通字段。
            continue
        # 单列表格：第 0 行是标题，第 1 行是长文本内容区（原有文字就是填写要求）
        if len(table.rows) >= 2 and len(table.rows[0].cells) == 1:
            label = table.rows[0].cells[0].text.strip()
            if is_valid_label(label):
                fields.append({
                    "label": label.rstrip("：:"),
                    "field_type": "long_text",
                    "confidence": 0.90,
                    "evidence": "单列表格中标题下方的空白内容区",
                    "instruction": table.rows[1].cells[0].text.strip(),
                    "location": {"file_type": "docx", "table_index": table_index, "row_index": 1, "cell_index": 0},
                    "needs_ai_generation": True,
                })
            continue
        # 表格顶部的标题行（整行只有一格内容）不参与填写
        title_rows = find_docx_title_rows(table)
        for row_index, row in enumerate(table.rows):
            if row_index in title_rows:
                continue
            for cell_index, cell in enumerate(row.cells):
                cell_tc = cell._tc
                checkbox_options = extract_checkbox_options(cell.text)
                if any(cell_tc is seen for seen in seen_cells) or (cell.text.strip() and not checkbox_options):
                    continue
                seen_cells.append(cell_tc)
                # 只认紧邻的标题，避免表格下方多余的空白行被重复识别成待填项。
                left_cell = row.cells[cell_index - 1] if cell_index > 0 else None
                above_cell = None
                if row_index > 0 and cell_index < len(table.rows[row_index - 1].cells):
                    above_cell = table.rows[row_index - 1].cells[cell_index]
                left = left_cell.text.strip() if left_cell is not None else ""
                above = above_cell.text.strip() if above_cell is not None else ""
                candidates = []
                if left:
                    candidates.append((left, left_cell, 0.93, "同一行左侧相邻单元格为标题"))
                if above:
                    candidates.append((above, above_cell, 0.82, "同一列上方相邻单元格为标题"))
                # 左侧是序号、示例文字时，自动改用上方的表头
                picked = next((item for item in candidates if is_valid_label(item[0])), None)
                if picked is None:
                    continue
                label, label_cell, confidence, evidence = picked
                # 合并标题会让多个空白格共用同一个标题，这里只保留最先遇到的一个。
                label_tc = label_cell._tc
                if any(label_tc is seen for seen in seen_labels):
                    continue
                seen_labels.append(label_tc)
                if checkbox_options:
                    fields.append({
                        "label": label.rstrip("：:"),
                        "field_type": "choice",
                        "confidence": confidence,
                        "evidence": "标题相邻单元格包含 Word 单选框选项",
                        "instruction": f"这是单选题，只能从以下候选项中选择一项：{'、'.join(checkbox_options)}。",
                        "location": {
                            "file_type": "docx", "table_index": table_index,
                            "row_index": row_index, "cell_index": cell_index,
                            "checkbox_options": checkbox_options,
                        },
                        "needs_ai_generation": False,
                    })
                    continue
                field_type, needs_ai_generation = infer_field_type(label)
                fields.append({
                    "label": label.rstrip("：:"),
                    "field_type": field_type,
                    "confidence": confidence,
                    "evidence": evidence,
                    "location": {"file_type": "docx", "table_index": table_index, "row_index": row_index, "cell_index": cell_index},
                    "needs_ai_generation": needs_ai_generation,
                })

    blank_pattern = re.compile(r"(?P<label>[^：:\n]{1,40})[：:]\s*(?P<blank>_{3,}|　{3,}|\s{4,})")
    for paragraph_index, paragraph in enumerate(document.paragraphs):
        for match in blank_pattern.finditer(paragraph.text):
            label = match.group("label").strip()
            if not is_valid_label(label):
                continue
            field_type, needs_ai_generation = infer_field_type(label)
            fields.append({
                "label": label,
                "field_type": field_type,
                "confidence": 0.75,
                "evidence": "段落中存在“标题：下划线/空白区域”模式",
                "location": {"file_type": "docx", "paragraph_index": paragraph_index, "character_start": match.start("blank"), "character_end": match.end("blank")},
                "needs_ai_generation": needs_ai_generation,
            })
    return fields


def detect_form_fields(file_path: Path) -> list[dict]:
    if file_path.suffix.lower() == ".xlsx":
        return detect_xlsx_fields(file_path)
    if file_path.suffix.lower() == ".docx":
        return detect_docx_fields(file_path)
    raise ValueError("仅支持 .docx 与 .xlsx 文件")


def write_xlsx_fields(source_path: Path, output_path: Path, fields: list[sqlite3.Row]) -> None:
    workbook = load_workbook(source_path)
    try:
        for field in fields:
            location = json.loads(field["location_json"])
            if location.get("file_type") != "xlsx":
                continue
            if field["field_type"] == "repeat_table":
                repeat_rows = normalize_repeat_rows(field["final_value"], location.get("columns", []))
                if not repeat_rows:
                    continue
                worksheet = workbook[location["sheet"]]
                data_start = location["data_start_row"]
                template_rows = location.get("template_row_indexes", [data_start])
                capacity = len(template_rows)
                # 将新增行插入“说明/备注”等尾部内容前，再完整复制最后一条模板行样式。
                extra_rows = max(0, len(repeat_rows) - capacity)
                if extra_rows:
                    insertion_row = template_rows[-1] + 1
                    source_row = template_rows[-1]
                    worksheet.insert_rows(insertion_row, amount=extra_rows)
                    for offset in range(extra_rows):
                        copy_xlsx_row_style(worksheet, source_row, insertion_row + offset, worksheet.max_column)
                for row_offset, row_values in enumerate(repeat_rows):
                    target_row = data_start + row_offset
                    for column in location.get("columns", []):
                        worksheet.cell(target_row, column["column"]).value = row_values.get(column["key"], "")
                    if serial_column := location.get("serial_column"):
                        worksheet.cell(target_row, serial_column).value = row_offset + 1
                continue
            if not field["final_value"]:
                continue
            workbook[location["sheet"]][location["cell"]] = field["final_value"]
        workbook.save(output_path)
    finally:
        workbook.close()


def write_docx_fields(source_path: Path, output_path: Path, fields: list[sqlite3.Row]) -> None:
    document = Document(source_path)
    paragraph_replacements: dict[int, list[tuple[int, int, str]]] = {}
    for field in fields:
        location = json.loads(field["location_json"])
        if location.get("file_type") != "docx":
            continue
        if field["field_type"] == "repeat_table":
            repeat_rows = normalize_repeat_rows(field["final_value"], location.get("columns", []))
            if not repeat_rows:
                continue
            table = document.tables[location["table_index"]]
            data_start = location["data_start_row"]
            template_indexes = location.get("template_row_indexes", [data_start])
            capacity = len(template_indexes)
            # 需求条数超过模板预留行时，深拷贝最后一行，保留边框、底色、行高等格式。
            last_row_element = table.rows[template_indexes[-1]]._tr
            for _ in range(max(0, len(repeat_rows) - capacity)):
                copied_row = deepcopy(last_row_element)
                last_row_element.addnext(copied_row)
                last_row_element = copied_row
            for row_offset, row_values in enumerate(repeat_rows):
                target_row = table.rows[data_start + row_offset]
                for column in location.get("columns", []):
                    cell_index = column["cell_index"]
                    if cell_index < len(target_row.cells):
                        target_row.cells[cell_index].text = row_values.get(column["key"], "")
            continue
        checkbox_options = location.get("checkbox_options")
        if checkbox_options:
            if field["final_value"]:
                cell = document.tables[location["table_index"]].rows[location["row_index"]].cells[location["cell_index"]]
                mark_docx_checkbox(cell, field["final_value"], checkbox_options)
            continue
        # 带“填写要求”的单元格即使没生成内容也要写入，用来清掉原文里的要求文字
        if not field["final_value"] and not field["instruction"]:
            continue
        if "table_index" in location:
            document.tables[location["table_index"]].rows[location["row_index"]].cells[location["cell_index"]].text = field["final_value"]
        elif "paragraph_index" in location:
            paragraph_replacements.setdefault(location["paragraph_index"], []).append(
                (location["character_start"], location["character_end"], field["final_value"])
            )
    # 逆序替换防止一个段落内前面的文本长度变化影响后续位置。
    for paragraph_index, replacements in paragraph_replacements.items():
        paragraph = document.paragraphs[paragraph_index]
        text = paragraph.text
        for start, end, value in sorted(replacements, reverse=True):
            text = f"{text[:start]}{value}{text[end:]}"
        paragraph.text = text
    document.save(output_path)


def export_filled_form(form: sqlite3.Row, fields: list[sqlite3.Row]) -> Path:
    source_path = Path(form["stored_path"])
    # 用用户原始文件名，而不是内部的表单 ID，否则导出名字会是一串 UUID
    original_stem = Path(form["original_name"]).stem
    output_name = f"{original_stem}_已填写_{datetime.now().strftime('%Y%m%d_%H%M%S')}{source_path.suffix.lower()}"
    output_path = current_outputs_dir() / output_name
    if source_path.suffix.lower() == ".xlsx":
        write_xlsx_fields(source_path, output_path, fields)
    elif source_path.suffix.lower() == ".docx":
        write_docx_fields(source_path, output_path, fields)
    else:
        raise ValueError("不支持导出此文件类型")
    return output_path


def serialize_record(record: sqlite3.Row, include_extracted_text: bool = False) -> dict:
    result = {
        "id": record["id"],
        "name": record["original_name"],
        "extension": record["extension"],
        "kind": {".docx": "Word 文档", ".xlsx": "Excel 表格", ".pdf": "PDF 文档"}.get(record["extension"], "文件"),
        "source_path": record["source_path"],
        "imported_at": record["imported_at"],
        "parse_status": record["parse_status"],
        "parse_error": record["parse_error"],
        "text_preview": record["extracted_text"][:500],
    }
    if include_extracted_text:
        result["extracted_text"] = record["extracted_text"]
    return result


def serialize_form(record: sqlite3.Row, fields: list[sqlite3.Row]) -> dict:
    return {
        "id": record["id"],
        "name": record["original_name"],
        "extension": record["extension"],
        "imported_at": record["imported_at"],
        "fields": [
            {
                "id": field["id"],
                "label": field["raw_label"],
                "field_type": field["field_type"],
                "confidence": field["confidence"],
                "evidence": field["evidence"],
                "location": json.loads(field["location_json"]),
                "current_value": field["current_value"],
                "needs_ai_generation": bool(field["needs_ai_generation"]),
                "suggested_value": field["suggested_value"],
                "suggestion_reason": field["suggestion_reason"],
                "source_ids": json.loads(field["source_ids_json"]),
                "generation_status": field["generation_status"],
                "final_value": field["final_value"],
                "instruction": field["instruction"],
            }
            for field in fields
        ],
    }


def token_set(text: str) -> set[str]:
    return set(re.findall(r"[a-zA-Z0-9@._-]+|[\u4e00-\u9fff]", text.lower()))


def retrieve_relevant_knowledge(connection: sqlite3.Connection, label: str, broad: bool = False) -> list[dict]:
    """挑选要发给模型的资料片段。

    先按关键词重叠挑最相关的片段，再给"每一份资料"补几条兜底：个人资料里
    字段名常常不会原样出现（资料写的是"陈雅雯，女，27岁"，字段却叫"姓名"），
    只靠字面重叠会一条都搜不到。长文本字段取更宽的取样范围。
    """
    label_tokens = token_set(label)
    limit = 15 if broad else 10
    scored: list[tuple[int, dict]] = []
    supplements: list[dict] = []
    records = connection.execute(
        "SELECT id, original_name, extracted_text FROM knowledge_files WHERE parse_status = 'parsed'"
    ).fetchall()
    for record in records:
        parts = [part.strip()[:1_000] for part in re.split(r"\n{1,}", record["extracted_text"])]
        parts = [part for part in parts if part]
        matched = [part for part in parts if label_tokens & token_set(part)]
        for part in matched:
            scored.append((len(label_tokens & token_set(part)), {
                "id": record["id"], "name": record["original_name"], "text": part,
            }))
        # 关键词完全不重合时，至少让模型看到这份资料的开头几条
        for part in [item for item in parts if item not in matched][:4]:
            supplements.append({"id": record["id"], "name": record["original_name"], "text": part})
    scored.sort(key=lambda item: item[0], reverse=True)
    result = [item for _, item in scored]
    for item in supplements:
        if len(result) >= limit:
            break
        result.append(item)
    return result[:limit]


PHONE_VALUE_PATTERN = re.compile(r"(?<!\d)(?:\+?86[-\s]?)?(1[3-9]\d{9})(?!\d)")
EMAIL_VALUE_PATTERN = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")


def canonical_basic_field(label: str) -> str | None:
    """将常见表单标题归一成可由规则可靠提取的基础信息类型。"""
    normalized = re.sub(r"\s+", "", label).lower().strip("：:")
    if any(item in normalized for item in ("联系电话", "联系手机", "手机号", "手机号码", "电话")):
        return "phone"
    if "邮箱" in normalized or "email" in normalized:
        return "email"
    if normalized in {"姓名", "名字", "申请人姓名", "填表人姓名"}:
        return "name"
    if normalized in {"性别", "本人性别"}:
        return "gender"
    if normalized in {"公司", "单位", "工作单位", "就职单位", "任职单位", "所在公司", "公司名称"}:
        return "company"
    if normalized in {"职业", "职位", "职务", "岗位", "当前职位"}:
        return "position"
    if any(item in normalized for item in ("家庭地址", "居住地址", "联系地址", "住址", "地址")):
        return "address"
    if any(item in normalized for item in ("最高学历", "学历")):
        return "education"
    return None


def reliable_rule_value(label: str, contexts: list[dict]) -> dict | None:
    """从资料片段提取唯一可信的基础信息；多个不同候选时交给 LLM 判断。"""
    field_kind = canonical_basic_field(label)
    if not field_kind:
        return None
    candidates: list[tuple[str, str]] = []
    for context in contexts:
        text = context["text"]
        values: list[str] = []
        if field_kind == "phone":
            values = [match.group(1) for match in PHONE_VALUE_PATTERN.finditer(text)]
        elif field_kind == "email":
            values = EMAIL_VALUE_PATTERN.findall(text)
        elif field_kind == "name":
            values = re.findall(r"(?:姓名|名字)\s*[:：]\s*([\u4e00-\u9fff]{2,5})", text)
            if not values:
                values = re.findall(r"(?:^|[\n。])\s*([\u4e00-\u9fff]{2,5})\s*[，,]\s*(?:男|女)", text)
        elif field_kind == "gender":
            values = re.findall(r"(?:性别\s*[:：]?\s*|[，,]\s*)(男|女)(?=[，,。；;\s])", text)
        elif field_kind == "company":
            values = re.findall(r"(?:工作单位|就职单位|任职单位|公司名称)\s*[:：]\s*([^\n，。；;]{2,80})", text)
            values += re.findall(r"就职于\s*([^\n，。；;]{2,80})", text)
        elif field_kind == "position":
            values = re.findall(r"(?:职业|职位|职务|岗位)\s*(?:为|是|[:：])\s*([^\n，。；;]{2,60})", text)
        elif field_kind == "address":
            values = re.findall(r"(?:家庭地址|居住地址|联系地址|住址|地址)\s*(?:是|为|[:：])\s*([^\n。；;]{2,120})", text)
        elif field_kind == "education":
            values = re.findall(r"(?:最高学历|学历)\s*(?:是|为|[:：])\s*([^\n，。；;]{2,20})", text)
        for value in values:
            cleaned = re.sub(r"\s+", " ", value).strip(" ，,。；;：:")
            if cleaned:
                candidates.append((cleaned, context["id"]))

    values_to_sources: dict[str, set[str]] = {}
    for value, source_id in candidates:
        values_to_sources.setdefault(value, set()).add(source_id)
    # 同一资料在多个文件重复出现是正常情况；只要值唯一就可由规则直接填写。
    if len(values_to_sources) != 1:
        return None
    value, source_ids = next(iter(values_to_sources.items()))
    return {
        "value": value,
        "reason": "已由本地资料中的明确基础信息自动匹配",
        "source_ids": sorted(source_ids),
    }


def request_llm_suggestions(
    fields: list[sqlite3.Row],
    contexts_by_field: dict[str, list[dict]],
    writing_style: str,
) -> dict[str, dict]:
    """规则优先填写基础信息，其余字段通过一次 LLM 请求兜底。"""
    result_by_field: dict[str, dict] = {}
    pending_fields: list[sqlite3.Row] = []
    for field in fields:
        rule_result = reliable_rule_value(field["raw_label"], contexts_by_field[field["id"]])
        if rule_result:
            result_by_field[field["id"]] = rule_result
        else:
            pending_fields.append(field)

    if not pending_fields:
        return result_by_field

    settings = load_settings()
    api_key = settings.get("api_key", "")
    model = settings.get("model", "")
    base_url = settings.get("base_url", "")
    if not api_key or not model or not base_url:
        raise HTTPException(
            status_code=503,
            detail="尚未在客户端完成模型地址、模型名称和密钥配置，无法生成 AI 建议。",
        )

    sources: list[dict] = []
    source_indexes: dict[tuple[str, str], int] = {}
    field_payloads: list[dict] = []
    for field in pending_fields:
        available_source_indexes: list[int] = []
        for context in contexts_by_field[field["id"]]:
            key = (context["id"], context["text"])
            if key not in source_indexes:
                source_indexes[key] = len(sources) + 1
                sources.append(context)
            available_source_indexes.append(source_indexes[key])
        field_payloads.append({
            "field_id": field["id"],
            "title": field["raw_label"],
            "type": field["field_type"],
            "needs_long_text": bool(field["needs_ai_generation"]),
            "instruction": field["instruction"] or "",
            "columns": json.loads(field["location_json"]).get("columns", []) if field["field_type"] == "repeat_table" else [],
            "available_source_indexes": available_source_indexes,
        })

    source_text = "\n\n".join(
        f"[资料 {index + 1}：{context['name']}]\n{context['text']}"
        for index, context in enumerate(sources)
    ) or "（资料库中没有检索到相关内容）"
    system_prompt = (
        "你是本地智能填表助手。只能基于提供的资料生成建议，不得编造事实。"
        "资料可能是不带字段标签的自然段落，请自行判断哪部分对应所需字段。"
        "每个字段只能引用该字段的 available_source_indexes。若资料不足，value 返回空字符串，并在 reason 说明缺少什么资料。"
        "若字段给出了字数或内容要求，必须严格遵守。"
        "必须返回 JSON 对象，且格式严格为 "
        "{\"suggestions\":[{\"field_id\":\"字段ID\",\"value\":\"...\",\"reason\":\"...\",\"source_indexes\":[1,2]}]}。"
        "当 type 为 repeat_table 时，value 必须是空字符串，改为返回 rows 数组；rows 的每一项是一个项目，"
        "对象键必须使用 columns 中的 key，每个项目只占一项，不能把多个项目拼进同一项。"
        "必须为每个 field_id 返回且只返回一条建议。"
    )
    user_prompt = (
        f"写作风格：{writing_style}\n\n"
        f"待填写字段：\n{json.dumps(field_payloads, ensure_ascii=False)}\n\n"
        f"可用资料：\n{source_text}"
    )
    base_url = base_url.rstrip("/")
    request_body = json.dumps({
        "model": model,
        "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": user_prompt}],
        "temperature": 0.2,
        "response_format": {"type": "json_object"},
    }).encode("utf-8")
    request = urllib.request.Request(
        f"{base_url}/chat/completions",
        data=request_body,
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=LLM_BATCH_TIMEOUT_SECONDS) as response:
            response_data = json.loads(response.read().decode("utf-8"))
    except TimeoutError as error:
        raise HTTPException(
            status_code=504,
            detail=f"模型服务在 {LLM_BATCH_TIMEOUT_SECONDS} 秒内未返回结果，请稍后重试或检查模型服务。",
        ) from error
    except urllib.error.HTTPError as error:
        detail = error.read().decode("utf-8", errors="replace")[:500]
        raise HTTPException(status_code=502, detail=f"LLM 服务返回错误：{detail}") from error
    except urllib.error.URLError as error:
        raise HTTPException(status_code=502, detail=f"无法连接 LLM 服务：{error.reason}") from error

    try:
        content = response_data["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as error:
        raise HTTPException(status_code=502, detail="LLM 返回结构不符合聊天补全接口格式") from error
    try:
        result = json.loads(content)
    except json.JSONDecodeError as error:
        raise HTTPException(status_code=502, detail="LLM 未返回有效 JSON 结果") from error
    raw_suggestions = result.get("suggestions", [])
    if not isinstance(raw_suggestions, list):
        raise HTTPException(status_code=502, detail="LLM 未返回字段建议列表")

    suggestions_by_id = {
        item.get("field_id"): item for item in raw_suggestions
        if isinstance(item, dict) and isinstance(item.get("field_id"), str)
    }
    for field in pending_fields:
        suggestion = suggestions_by_id.get(field["id"])
        if not suggestion:
            result_by_field[field["id"]] = {
                "value": "",
                "reason": "模型未返回该字段的建议，请人工填写。",
                "source_ids": [],
            }
            continue
        allowed_source_ids = {context["id"] for context in contexts_by_field[field["id"]]}
        source_ids = [
            sources[index - 1]["id"] for index in suggestion.get("source_indexes", [])
            if isinstance(index, int) and 1 <= index <= len(sources)
            and sources[index - 1]["id"] in allowed_source_ids
        ]
        value = str(suggestion.get("value", "")).strip()
        reason = str(suggestion.get("reason", "")).strip()
        if field["field_type"] == "repeat_table":
            try:
                columns = json.loads(field["location_json"]).get("columns", [])
            except (TypeError, json.JSONDecodeError):
                columns = []
            rows = normalize_repeat_rows(suggestion.get("rows", []), columns)
            value = json.dumps(rows, ensure_ascii=False) if rows else ""
            if not rows and not reason:
                reason = "资料中未提取到可填写的项目经历，请人工补充。"
        if field["field_type"] == "choice":
            try:
                choice_options = json.loads(field["location_json"]).get("checkbox_options", [])
            except (TypeError, json.JSONDecodeError):
                choice_options = []
            normalized_choice = normalize_checkbox_choice(value, choice_options)
            if normalized_choice:
                value = normalized_choice
            else:
                value = ""
                reason = f"请从候选项中选择一项：{'、'.join(choice_options)}"
        result_by_field[field["id"]] = {
            "value": value,
            "reason": reason,
            "source_ids": list(dict.fromkeys(source_ids)),
        }
    return result_by_field


class ImportKnowledgeRequest(BaseModel):
    source_path: str = Field(min_length=1, description="由 Electron 文件选择框返回的绝对路径")
    replace_file_id: str | None = Field(default=None, description="用户确认更新的同名资料 ID")


class ImportFormRequest(BaseModel):
    source_path: str = Field(min_length=1, description="由 Electron 文件选择框返回的绝对路径")


class GenerateSuggestionsRequest(BaseModel):
    # 不传时使用「AI 设置」里保存的写作风格
    writing_style: str | None = Field(default=None, max_length=2_000)


class LlmSettingsUpdate(BaseModel):
    base_url: str = Field(min_length=1, max_length=500)
    model: str = Field(min_length=1, max_length=200)
    api_key: str | None = Field(default=None, max_length=2_000)


class LlmConnectionTestRequest(BaseModel):
    """测试当前页面的配置；空密钥代表继续使用本机已保存的密钥。"""
    base_url: str = Field(min_length=1, max_length=500)
    model: str = Field(min_length=1, max_length=200)
    api_key: str | None = Field(default=None, max_length=2_000)


class WritingStyleItem(BaseModel):
    id: str = Field(min_length=1, max_length=100)
    name: str = Field(min_length=1, max_length=100)
    content: str = Field(min_length=1, max_length=2_000)


class WritingStylesUpdate(BaseModel):
    styles: list[WritingStyleItem] = Field(min_length=1)
    active_id: str | None = Field(default=None, max_length=100)


class StorageSettingsUpdate(BaseModel):
    output_dir: str | None = Field(default=None, max_length=1_000)


class NotificationSettingsUpdate(BaseModel):
    enabled: bool = True


class ExportPathCheck(BaseModel):
    path: str = Field(min_length=1, max_length=1_000)


class FieldValue(BaseModel):
    field_id: str
    value: str = Field(default="", max_length=20_000)


class ExportFormRequest(BaseModel):
    values: list[FieldValue]


app = FastAPI(title="智填助手本地服务", version="0.1.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.on_event("startup")
def startup() -> None:
    ensure_storage()


@app.get("/health")
def health_check() -> dict:
    return {"status": "ok", "data_root": str(DATA_ROOT)}


@app.get("/api/system/paths")
def get_storage_paths() -> dict:
    ensure_storage()
    return {
        "data_root": str(DATA_ROOT),
        "knowledge": str(KNOWLEDGE_DIR),
        "forms": str(FORMS_DIR),
        "outputs": str(current_outputs_dir()),
        "default_outputs": str(OUTPUTS_DIR),
    }


@app.get("/api/settings/llm")
def get_llm_settings() -> dict:
    settings = load_settings()
    return {
        "base_url": settings.get("base_url", ""),
        "model": settings.get("model", ""),
        "has_api_key": bool(settings.get("api_key")),
    }


@app.post("/api/settings/llm/test")
def test_llm_connection(payload: LlmConnectionTestRequest) -> dict:
    """以一次极小的 Chat Completions 请求验证当前输入的模型配置。

    此接口不会调用 save_settings；用户可以先测试尚未保存的新地址、模型或密钥。
    """
    previous_settings = load_settings()
    base_url = payload.base_url.strip().rstrip("/")
    model = payload.model.strip()
    api_key = payload.api_key.strip() if payload.api_key and payload.api_key.strip() else previous_settings.get("api_key", "")
    if not base_url or not model:
        raise HTTPException(status_code=422, detail="请填写模型服务地址和模型名称")
    if not api_key:
        raise HTTPException(status_code=422, detail="请填写 API 密钥，或先保存已有密钥")

    request_body = json.dumps({
        "model": model,
        "messages": [{"role": "user", "content": "请仅回复 OK"}],
        "temperature": 0,
        "max_tokens": 8,
    }).encode("utf-8")
    request = urllib.request.Request(
        f"{base_url}/chat/completions",
        data=request_body,
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        method="POST",
    )
    started_at = time.perf_counter()
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            response_data = json.loads(response.read().decode("utf-8"))
    except TimeoutError as error:
        raise HTTPException(status_code=504, detail="连接模型服务超时（20 秒），请检查网络或服务地址") from error
    except urllib.error.HTTPError as error:
        provider_message = error.read().decode("utf-8", errors="replace")[:300]
        if error.code in (401, 403):
            detail = "API 密钥无效、已过期，或没有调用该模型的权限"
        elif error.code == 404:
            detail = "接口地址或模型名称不存在；请确认地址是 OpenAI Chat Completions 的 API 根地址"
        elif error.code == 429:
            detail = "模型服务请求过于频繁或额度不足"
        else:
            detail = f"模型服务返回 HTTP {error.code}"
        if provider_message:
            detail = f"{detail}：{provider_message}"
        raise HTTPException(status_code=502, detail=detail) from error
    except urllib.error.URLError as error:
        raise HTTPException(status_code=502, detail=f"无法连接模型服务：{error.reason}") from error
    except (OSError, json.JSONDecodeError) as error:
        raise HTTPException(status_code=502, detail=f"模型服务响应异常：{error}") from error

    try:
        response_data["choices"][0]
    except (KeyError, IndexError, TypeError) as error:
        raise HTTPException(status_code=502, detail="模型服务返回结构不符合 Chat Completions 接口格式") from error
    return {
        "success": True,
        "model": response_data.get("model") or model,
        "latency_ms": round((time.perf_counter() - started_at) * 1000),
    }


@app.put("/api/settings/llm")
def update_llm_settings(payload: LlmSettingsUpdate) -> dict:
    previous_settings = load_settings()
    api_key = payload.api_key.strip() if payload.api_key else previous_settings.get("api_key", "")
    if not api_key:
        raise HTTPException(status_code=422, detail="请填写 API 密钥")
    # 合并写入：不能用新字典整体覆盖，否则 settings.json 里其它配置（如 output_dir）会被抹掉
    settings = previous_settings
    settings.update({
        "base_url": payload.base_url.rstrip("/"),
        "model": payload.model.strip(),
        "api_key": api_key,
    })
    save_settings(settings)
    return {
        "base_url": payload.base_url.rstrip("/"),
        "model": payload.model.strip(),
        "has_api_key": True,
    }


@app.get("/api/settings/writing-styles")
def get_writing_styles() -> dict:
    """写作风格列表和当前选中的风格。"""
    return load_writing_styles()


@app.put("/api/settings/writing-styles")
def update_writing_styles(payload: WritingStylesUpdate) -> dict:
    """整体保存写作风格列表；active_id 不在列表里时自动落到第一条。"""
    settings = load_settings()
    settings["writing_styles"] = [item.model_dump() for item in payload.styles]
    active_id = payload.active_id if any(item.id == payload.active_id for item in payload.styles) else payload.styles[0].id
    settings["active_style_id"] = active_id
    save_settings(settings)
    return load_writing_styles()


@app.get("/api/settings/storage")
def get_storage_settings() -> dict:
    """导出目录配置。effective 是当前真正生效的位置。"""
    settings = load_settings()
    return {
        "output_dir": settings.get("output_dir", ""),
        "default_output_dir": str(OUTPUTS_DIR),
        "effective_output_dir": str(current_outputs_dir()),
    }


@app.put("/api/settings/storage")
def update_storage_settings(payload: StorageSettingsUpdate) -> dict:
    """保存导出目录；传空字符串表示恢复默认目录。"""
    output_dir = (payload.output_dir or "").strip()
    if output_dir:
        target = Path(output_dir).expanduser()
        try:
            target.mkdir(parents=True, exist_ok=True)
        except OSError as error:
            raise HTTPException(status_code=422, detail=f"无法使用该目录：{error}") from error
        if not target.is_dir():
            raise HTTPException(status_code=422, detail="该路径不是有效的文件夹")

    settings = load_settings()
    settings["output_dir"] = output_dir
    save_settings(settings)
    return {
        "output_dir": output_dir,
        "default_output_dir": str(OUTPUTS_DIR),
        "effective_output_dir": str(current_outputs_dir()),
    }


@app.get("/api/settings/notifications")
def get_notification_settings() -> dict:
    """填写完成后的桌面通知开关；未配置时默认开启。"""
    settings = load_settings()
    return {"enabled": settings.get("notifications_enabled", True)}


@app.put("/api/settings/notifications")
def update_notification_settings(payload: NotificationSettingsUpdate) -> dict:
    settings = load_settings()
    settings["notifications_enabled"] = payload.enabled
    save_settings(settings)
    return {"enabled": payload.enabled}


@app.get("/api/knowledge")
def list_knowledge_files() -> list[dict]:
    with database_connection() as connection:
        records = connection.execute("SELECT * FROM knowledge_files ORDER BY imported_at DESC").fetchall()
    return [serialize_record(record) for record in records]


@app.get("/api/knowledge/{file_id}")
def get_knowledge_file(file_id: str) -> dict:
    """返回一份资料的完整解析内容，供客户端预览。"""
    with database_connection() as connection:
        record = connection.execute(
            "SELECT * FROM knowledge_files WHERE id = ?", (file_id,)
        ).fetchone()
    if not record:
        raise HTTPException(status_code=404, detail="资料记录不存在")
    return serialize_record(record, include_extracted_text=True)


@app.post("/api/knowledge/import/check")
def check_knowledge_import(payload: ImportKnowledgeRequest) -> dict:
    """检查同名但不同来源路径的资料，供前端在真正导入前让用户确认。"""
    source_path = Path(payload.source_path).expanduser()
    if not source_path.is_file():
        raise HTTPException(status_code=404, detail="找不到要导入的文件")
    if source_path.suffix.lower() not in KNOWLEDGE_EXTENSIONS:
        raise HTTPException(status_code=400, detail="仅支持 Word（.docx）、Excel（.xlsx）和文本型 PDF（.pdf）文件")

    source_path = source_path.resolve()
    with database_connection() as connection:
        candidates = connection.execute(
            """SELECT * FROM knowledge_files
            WHERE original_name = ? AND source_path != ?
            ORDER BY imported_at DESC""",
            (source_path.name, str(source_path)),
        ).fetchall()
    return {"candidates": [serialize_record(record) for record in candidates]}


@app.post("/api/knowledge/import", status_code=status.HTTP_201_CREATED)
def import_knowledge_file(payload: ImportKnowledgeRequest) -> dict:
    source_path = Path(payload.source_path).expanduser()
    if not source_path.is_file():
        raise HTTPException(status_code=404, detail="找不到要导入的文件")
    if source_path.suffix.lower() not in KNOWLEDGE_EXTENSIONS:
        raise HTTPException(status_code=400, detail="仅支持 Word（.docx）、Excel（.xlsx）和文本型 PDF（.pdf）文件")

    source_path = source_path.resolve()
    file_hash = sha256_for_file(source_path)
    with database_connection() as connection:
        same_path_record = connection.execute(
            "SELECT * FROM knowledge_files WHERE source_path = ? ORDER BY imported_at DESC LIMIT 1",
            (str(source_path),),
        ).fetchone()
        target_record = same_path_record
        if not target_record and payload.replace_file_id:
            target_record = connection.execute(
                "SELECT * FROM knowledge_files WHERE id = ?", (payload.replace_file_id,)
            ).fetchone()
            if not target_record:
                raise HTTPException(status_code=404, detail="要更新的资料不存在")
            if target_record["original_name"] != source_path.name:
                raise HTTPException(status_code=400, detail="只能更新同名资料")

        if target_record:
            if target_record["sha256"] == file_hash:
                if target_record["source_path"] != str(source_path):
                    connection.execute(
                        "UPDATE knowledge_files SET source_path = ?, imported_at = ? WHERE id = ?",
                        (str(source_path), datetime.now(UTC).isoformat(), target_record["id"]),
                    )
                    target_record = connection.execute(
                        "SELECT * FROM knowledge_files WHERE id = ?", (target_record["id"],)
                    ).fetchone()
                    result = serialize_record(target_record)
                    result["import_action"] = "updated"
                    return result
                result = serialize_record(target_record)
                result["import_action"] = "unchanged"
                return result

            same_content_record = connection.execute(
                "SELECT id FROM knowledge_files WHERE sha256 = ? AND id != ?",
                (file_hash, target_record["id"]),
            ).fetchone()
            if same_content_record:
                raise HTTPException(status_code=409, detail="该文件内容已作为另一份资料导入，请先在资料库中确认后再处理。")

            stored_path = Path(target_record["stored_path"])
            shutil.copy2(source_path, stored_path)
            parse_status, parse_error, extracted_text = "parsed", None, ""
            try:
                extracted_text = extract_text(stored_path)
            except Exception as error:  # 新版本仍保留，允许用户查看并重新导入
                parse_status, parse_error = "failed", str(error)

            connection.execute(
                """UPDATE knowledge_files
                SET original_name = ?, extension = ?, source_path = ?, sha256 = ?, imported_at = ?, parse_status = ?, parse_error = ?, extracted_text = ?
                WHERE id = ?""",
                (
                    source_path.name,
                    source_path.suffix.lower(),
                    str(source_path),
                    file_hash,
                    datetime.now(UTC).isoformat(),
                    parse_status,
                    parse_error,
                    extracted_text,
                    target_record["id"],
                ),
            )
            record = connection.execute(
                "SELECT * FROM knowledge_files WHERE id = ?", (target_record["id"],)
            ).fetchone()
            result = serialize_record(record)
            result["import_action"] = "updated"
            return result

        if existing := connection.execute("SELECT * FROM knowledge_files WHERE sha256 = ?", (file_hash,)).fetchone():
            result = serialize_record(existing)
            result["import_action"] = "unchanged"
            return result

        file_id = str(uuid.uuid4())
        stored_path = KNOWLEDGE_DIR / f"{file_id}{source_path.suffix.lower()}"
        shutil.copy2(source_path, stored_path)
        parse_status, parse_error, extracted_text = "parsed", None, ""
        try:
            extracted_text = extract_text(stored_path)
        except Exception as error:  # 保留原文件，允许后续重试解析
            parse_status, parse_error = "failed", str(error)

        connection.execute(
            """INSERT INTO knowledge_files
            (id, original_name, extension, source_path, stored_path, sha256, imported_at, parse_status, parse_error, extracted_text)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (file_id, source_path.name, source_path.suffix.lower(), str(source_path), str(stored_path), file_hash,
             datetime.now(UTC).isoformat(), parse_status, parse_error, extracted_text),
        )
        record = connection.execute("SELECT * FROM knowledge_files WHERE id = ?", (file_id,)).fetchone()
    result = serialize_record(record)
    result["import_action"] = "created"
    return result


@app.post("/api/knowledge/{file_id}/ocr")
def ocr_knowledge_pdf(file_id: str) -> dict:
    """用户确认后识别一份扫描版 PDF。

    PaddleOCR 会在首次调用时从其官方模型源下载模型；后续均使用本地缓存。
    """
    with database_connection() as connection:
        record = connection.execute(
            "SELECT * FROM knowledge_files WHERE id = ?", (file_id,)
        ).fetchone()
    if not record:
        raise HTTPException(status_code=404, detail="资料记录不存在")
    if record["extension"].lower() != ".pdf":
        raise HTTPException(status_code=400, detail="OCR 识别仅支持 PDF 资料")

    stored_path = Path(record["stored_path"])
    if not stored_path.is_file():
        raise HTTPException(status_code=404, detail="本地资料副本不存在，请重新导入文件")

    set_ocr_progress(file_id, "running", 5, "正在准备扫描 PDF 识别…")
    try:
        extracted_text = extract_scanned_pdf_text(
            stored_path,
            lambda progress, message: set_ocr_progress(file_id, "running", progress, message),
        )
    except HTTPException as error:
        set_ocr_progress(file_id, "failed", 0, error.detail)
        raise
    except Exception as error:
        with database_connection() as connection:
            connection.execute(
                "UPDATE knowledge_files SET parse_status = ?, parse_error = ? WHERE id = ?",
                ("failed", str(error), file_id),
            )
        set_ocr_progress(file_id, "failed", 0, f"扫描 PDF 识别失败：{error}")
        raise HTTPException(status_code=422, detail=f"扫描 PDF 识别失败：{error}") from error

    with database_connection() as connection:
        connection.execute(
            """UPDATE knowledge_files
            SET parse_status = ?, parse_error = NULL, extracted_text = ?, imported_at = ?
            WHERE id = ?""",
            ("parsed", extracted_text, datetime.now(UTC).isoformat(), file_id),
        )
        updated_record = connection.execute(
            "SELECT * FROM knowledge_files WHERE id = ?", (file_id,)
        ).fetchone()
    set_ocr_progress(file_id, "completed", 100, "OCR 识别完成")
    return serialize_record(updated_record, include_extracted_text=True)


@app.get("/api/knowledge/{file_id}/ocr-status")
def get_knowledge_ocr_status(file_id: str) -> dict:
    """返回当前 OCR 阶段，供桌面端显示不确定时长的模型下载/识别进度。"""
    return OCR_PROGRESS.get(file_id, {"state": "idle", "progress": 0, "message": "等待开始 OCR 识别"})


@app.delete("/api/knowledge/{file_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_knowledge_file(file_id: str) -> None:
    with database_connection() as connection:
        record = connection.execute("SELECT * FROM knowledge_files WHERE id = ?", (file_id,)).fetchone()
        if not record:
            raise HTTPException(status_code=404, detail="资料记录不存在")
        stored_path = Path(record["stored_path"])
        if stored_path.exists():
            stored_path.unlink()
        connection.execute("DELETE FROM knowledge_files WHERE id = ?", (file_id,))


@app.post("/api/forms/import", status_code=status.HTTP_201_CREATED)
def import_form_file(payload: ImportFormRequest) -> dict:
    source_path = Path(payload.source_path).expanduser()
    if not source_path.is_file():
        raise HTTPException(status_code=404, detail="找不到要分析的表单")
    if source_path.suffix.lower() not in FORM_EXTENSIONS:
        raise HTTPException(status_code=400, detail="仅支持 Word（.docx）和 Excel（.xlsx）文件")

    source_path = source_path.resolve()
    file_hash = sha256_for_file(source_path)
    with database_connection() as connection:
        existing = connection.execute("SELECT * FROM forms WHERE sha256 = ?", (file_hash,)).fetchone()
        if existing:
            if existing["detector_version"] < FORM_DETECTOR_VERSION:
                try:
                    detected_fields = detect_form_fields(Path(existing["stored_path"]))
                except Exception as error:
                    raise HTTPException(status_code=422, detail=f"无法重新解析表单：{error}") from error
                connection.execute("DELETE FROM form_fields WHERE form_id = ?", (existing["id"],))
                for field in detected_fields:
                    connection.execute(
                        """INSERT INTO form_fields
                        (id, form_id, raw_label, field_type, confidence, evidence, location_json, needs_ai_generation, instruction)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                        (str(uuid.uuid4()), existing["id"], field["label"], field["field_type"], field["confidence"],
                         field["evidence"], json.dumps(field["location"], ensure_ascii=False), int(field["needs_ai_generation"]),
                         field.get("instruction", "")),
                    )
                connection.execute(
                    "UPDATE forms SET detector_version = ? WHERE id = ?",
                    (FORM_DETECTOR_VERSION, existing["id"]),
                )
                existing = connection.execute("SELECT * FROM forms WHERE id = ?", (existing["id"],)).fetchone()
            fields = connection.execute("SELECT * FROM form_fields WHERE form_id = ?", (existing["id"],)).fetchall()
            return serialize_form(existing, fields)

        form_id = str(uuid.uuid4())
        stored_path = FORMS_DIR / f"{form_id}{source_path.suffix.lower()}"
        shutil.copy2(source_path, stored_path)
        try:
            detected_fields = detect_form_fields(stored_path)
        except Exception as error:
            stored_path.unlink(missing_ok=True)
            raise HTTPException(status_code=422, detail=f"无法解析表单：{error}") from error

        imported_at = datetime.now(UTC).isoformat()
        connection.execute(
            """INSERT INTO forms (id, original_name, extension, source_path, stored_path, sha256, imported_at, detector_version)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (form_id, source_path.name, source_path.suffix.lower(), str(source_path), str(stored_path), file_hash, imported_at, FORM_DETECTOR_VERSION),
        )
        for field in detected_fields:
            connection.execute(
                """INSERT INTO form_fields
                (id, form_id, raw_label, field_type, confidence, evidence, location_json, needs_ai_generation, instruction)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (str(uuid.uuid4()), form_id, field["label"], field["field_type"], field["confidence"],
                 field["evidence"], json.dumps(field["location"], ensure_ascii=False), int(field["needs_ai_generation"]),
                 field.get("instruction", "")),
            )
        form = connection.execute("SELECT * FROM forms WHERE id = ?", (form_id,)).fetchone()
        fields = connection.execute("SELECT * FROM form_fields WHERE form_id = ?", (form_id,)).fetchall()
    return serialize_form(form, fields)


@app.get("/api/forms/{form_id}")
def get_form(form_id: str) -> dict:
    with database_connection() as connection:
        form = connection.execute("SELECT * FROM forms WHERE id = ?", (form_id,)).fetchone()
        if not form:
            raise HTTPException(status_code=404, detail="表单记录不存在")
        fields = connection.execute("SELECT * FROM form_fields WHERE form_id = ?", (form_id,)).fetchall()
    return serialize_form(form, fields)


@app.post("/api/forms/{form_id}/suggestions")
def generate_form_suggestions(form_id: str, payload: GenerateSuggestionsRequest) -> dict:
    """检索本地资料库后，为每个已识别字段生成可审核的 AI 草稿。"""
    with database_connection() as connection:
        form = connection.execute("SELECT * FROM forms WHERE id = ?", (form_id,)).fetchone()
        if not form:
            raise HTTPException(status_code=404, detail="表单记录不存在")
        fields = connection.execute("SELECT * FROM form_fields WHERE form_id = ?", (form_id,)).fetchall()
        if not fields:
            raise HTTPException(status_code=422, detail="该表单没有可生成建议的识别字段")

        writing_style = payload.writing_style or current_writing_style()
        contexts_by_field = {
            field["id"]: retrieve_relevant_knowledge(
                connection, field["raw_label"], broad=bool(field["needs_ai_generation"])
            )
            for field in fields
        }
        suggestions = request_llm_suggestions(fields, contexts_by_field, writing_style)
        for field in fields:
            suggestion = suggestions[field["id"]]
            connection.execute(
                """UPDATE form_fields
                SET suggested_value = ?, suggestion_reason = ?, source_ids_json = ?, generation_status = 'generated'
                WHERE id = ?""",
                (suggestion["value"], suggestion["reason"], json.dumps(suggestion["source_ids"]), field["id"]),
            )
        updated_fields = connection.execute("SELECT * FROM form_fields WHERE form_id = ?", (form_id,)).fetchall()
    return serialize_form(form, updated_fields)


@app.post("/api/forms/{form_id}/export")
def export_form(form_id: str, payload: ExportFormRequest) -> dict:
    """保存用户确认值，并生成不覆盖原文件的填写结果。"""
    with database_connection() as connection:
        form = connection.execute("SELECT * FROM forms WHERE id = ?", (form_id,)).fetchone()
        if not form:
            raise HTTPException(status_code=404, detail="表单记录不存在")
        form_fields = connection.execute("SELECT * FROM form_fields WHERE form_id = ?", (form_id,)).fetchall()
        field_ids = {field["id"] for field in form_fields}
        for value in payload.values:
            if value.field_id not in field_ids:
                raise HTTPException(status_code=422, detail="存在不属于当前表单的字段")
            connection.execute("UPDATE form_fields SET final_value = ? WHERE id = ?", (value.value, value.field_id))
        updated_fields = connection.execute("SELECT * FROM form_fields WHERE form_id = ?", (form_id,)).fetchall()
        try:
            output_path = export_filled_form(form, updated_fields)
        except Exception as error:
            raise HTTPException(status_code=422, detail=f"无法导出填写结果：{error}") from error

        filled_count = sum(bool(field["final_value"]) for field in updated_fields)
        connection.execute(
            """INSERT INTO exports
            (id, form_id, form_name, output_path, output_name, filled_field_count, missing_field_count, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (str(uuid.uuid4()), form_id, form["original_name"], str(output_path), output_path.name,
             filled_count, len(updated_fields) - filled_count, datetime.now(UTC).isoformat()),
        )
    return {"output_path": str(output_path), "output_name": output_path.name, "filled_field_count": filled_count}


@app.get("/api/exports")
def list_exports() -> list[dict]:
    """返回历次导出记录；文件是否还在会一起返回，供客户端置灰。"""
    with database_connection() as connection:
        records = connection.execute("SELECT * FROM exports ORDER BY created_at DESC").fetchall()
    return [
        {
            "id": record["id"],
            "form_id": record["form_id"],
            "form_name": record["form_name"],
            "output_path": record["output_path"],
            "output_name": record["output_name"],
            "filled_field_count": record["filled_field_count"],
            "missing_field_count": record["missing_field_count"],
            "created_at": record["created_at"],
            "file_exists": Path(record["output_path"]).is_file(),
        }
        for record in records
    ]


@app.delete("/api/exports/{export_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_export(export_id: str) -> None:
    """只删除记录，不动磁盘上已经导出的文件。"""
    with database_connection() as connection:
        record = connection.execute("SELECT * FROM exports WHERE id = ?", (export_id,)).fetchone()
        if not record:
            raise HTTPException(status_code=404, detail="导出记录不存在")
        connection.execute("DELETE FROM exports WHERE id = ?", (export_id,))


@app.delete("/api/exports")
def clear_exports() -> dict:
    """清空应用内的填写记录，不删除任何已导出的磁盘文件。"""
    with database_connection() as connection:
        deleted_count = connection.execute("SELECT COUNT(*) FROM exports").fetchone()[0]
        connection.execute("DELETE FROM exports")
    return {"deleted_count": deleted_count}


@app.post("/api/exports/verify")
def verify_export_path(payload: ExportPathCheck) -> dict:
    """校验某个路径是不是本程序登记过的导出文件。

    客户端打开文件前会先问这里：用户改过导出目录之后，历史记录里指向旧目录的
    文件仍然允许打开。
    """
    target = str(Path(payload.path).expanduser().resolve())
    with database_connection() as connection:
        record = connection.execute("SELECT 1 FROM exports WHERE output_path = ?", (target,)).fetchone()
    return {"allowed": record is not None}


if __name__ == "__main__":
    # PyInstaller 环境中不能使用 reload：它会额外派生开发监控进程，
    # 而且无法按源码模块路径重新导入。生产版直接运行应用对象即可。
    uvicorn.run(app, host="127.0.0.1", port=8999, reload=False)
