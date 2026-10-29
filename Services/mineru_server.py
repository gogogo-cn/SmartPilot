# -*- coding: utf-8 -*-
"""
MinerU 文档解析 HTTP 服务
包装 MinerU (magic-pdf) 为 HTTP API，供 VB.NET 插件调用
支持：PDF 文本提取、图片提取、表格提取、公式识别（LaTeX）

启动方式：
  python mineru_server.py --port 8036

安装依赖：
  pip install "mineru[core]" fastapi uvicorn python-multipart
"""

import argparse
import base64
import json
import logging
import os
import re
import shutil
import sys
import tempfile
import time
import traceback
from pathlib import Path
from typing import Optional

# FastAPI 相关
try:
    from fastapi import FastAPI, UploadFile, File, Form, HTTPException
    from fastapi.responses import JSONResponse
    from fastapi.middleware.cors import CORSMiddleware
    import uvicorn
except ImportError:
    print("[MinerU Server] 缺少 fastapi/uvicorn，请执行: pip install fastapi uvicorn python-multipart")
    sys.exit(1)

# MinerU 核心
MINERU_AVAILABLE = False
MINERU_VERSION = ""
MINERU_ERROR = ""

try:
    import magic_pdf
    MINERU_AVAILABLE = True
    try:
        MINERU_VERSION = magic_pdf.__version__
    except Exception:
        MINERU_VERSION = "unknown"
except ImportError:
    MINERU_ERROR = "magic-pdf 未安装，请执行: pip install mineru"

# 日志配置
logging.basicConfig(
    level=logging.INFO,
    format="[MinerU Server] %(asctime)s %(levelname)s: %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger("mineru_server")

# ==================== FastAPI 应用 ====================

app = FastAPI(title="MinerU Document Parser", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ==================== 数据模型 ====================

class ParseResult:
    """解析结果封装"""
    def __init__(self):
        self.success = False
        self.full_text = ""
        self.markdown = ""
        self.images = []       # [{index, base64, alt_text, width, height}]
        self.tables = []       # [{index, html, markdown, row_count, col_count, position}]
        self.formulas = []     # [{index, latex, type, position}]
        self.pages = []        # [{page_number, text, tables, images, formulas}]
        self.metadata = {}     # {page_count, file_size, parse_time_ms}
        self.error = ""


# ==================== MinerU 解析核心 ====================

def parse_pdf_mineru(pdf_bytes: bytes, extract_images: bool = True) -> ParseResult:
    """
    使用 MinerU 解析 PDF
    提取：纯文本、Markdown、图片、表格(HTML)、公式(LaTeX)
    """
    result = ParseResult()
    start_time = time.time()

    if not MINERU_AVAILABLE:
        result.error = MINERU_ERROR
        return result

    temp_dir = None
    try:
        # 创建临时工作目录
        temp_dir = tempfile.mkdtemp(prefix="mineru_")
        image_dir = os.path.join(temp_dir, "images")
        os.makedirs(image_dir, exist_ok=True)

        # MinerU 2.0 API 调用
        from magic_pdf.data.data_reader_writer import FileBasedDataWriter, FileBasedDataReader
        from magic_pdf.data.dataset import PymuDocDataset

        # 读取 PDF 字节
        logger.info(f"开始解析 PDF, 大小: {len(pdf_bytes)} 字节")

        # 创建数据集
        ds = PymuDocDataset(pdf_bytes)

        # 执行解析（含公式检测 + 表格识别）
        parse_result = ds.apply(
            image_dir=image_dir,
            start_page_id=0,
            end_page_id=ds.page_count - 1,
            f_dump_md=True,
            f_dump_middle_json=True,
            f_dump_model_output=True,
            f_dump_orig_pdf=False,
            f_dump_content_list=True,
        )

        # 提取 Markdown 全文
        md_content = ""
        try:
            md_content = ds.get_markdown(image_dir)
        except Exception as e:
            logger.warning(f"获取 Markdown 失败: {e}")

        # 提取 content_list（结构化内容）
        content_list = []
        try:
            content_list = ds.get_content_list(image_dir)
        except Exception as e:
            logger.warning(f"获取 content_list 失败: {e}")

        # 提取 middle_json（含模型原始输出）
        middle_json = {}
        try:
            middle_json = ds.get_middle_json(image_dir)
        except Exception as e:
            logger.warning(f"获取 middle_json 失败: {e}")

        # 解析 content_list 提取表格、图片、公式
        tables = []
        images = []
        formulas = []
        page_texts = []
        current_page = 0
        page_text_builder = []

        if content_list:
            for item in content_list:
                item_type = item.get("type", "")

                if item_type == "text":
                    text = item.get("text", "")
                    if text:
                        page_text_builder.append(text)

                elif item_type == "table":
                    table_info = {
                        "index": len(tables),
                        "html": item.get("html", ""),
                        "markdown": item.get("text", item.get("markdown", "")),
                        "row_count": item.get("row_count", 0),
                        "col_count": item.get("col_count", 0),
                        "page": item.get("page_idx", 0),
                        "position": item.get("bbox", [])
                    }
                    tables.append(table_info)
                    # 表格也加入文本
                    if table_info["markdown"]:
                        page_text_builder.append(f"\n[表格]\n{table_info['markdown']}\n")

                elif item_type == "image":
                    if extract_images:
                        img_path = item.get("image_path", "")
                        img_base64 = ""
                        if img_path and os.path.exists(img_path):
                            with open(img_path, "rb") as img_f:
                                img_base64 = base64.b64encode(img_f.read()).decode("utf-8")
                        img_info = {
                            "index": len(images),
                            "base64": img_base64,
                            "alt_text": item.get("alt_text", item.get("text", "")),
                            "width": item.get("width", 0),
                            "height": item.get("height", 0),
                            "page": item.get("page_idx", 0)
                        }
                        images.append(img_info)

                elif item_type == "formula":
                    formula_info = {
                        "index": len(formulas),
                        "latex": item.get("text", item.get("latex", "")),
                        "type": item.get("formula_type", "unknown"),  # inline / display
                        "page": item.get("page_idx", 0),
                        "position": item.get("bbox", [])
                    }
                    formulas.append(formula_info)
                    # 公式以 LaTeX 形式加入文本
                    if formula_info["latex"]:
                        latex = formula_info["latex"]
                        if formula_info["type"] == "display":
                            page_text_builder.append(f"\n$$\n{latex}\n$$\n")
                        else:
                            page_text_builder.append(f"${latex}$")

        # 拼接全文
        full_text = md_content if md_content else "\n".join(page_text_builder)

        # 构建分页信息
        if content_list:
            page_map = {}
            for item in content_list:
                pg = item.get("page_idx", 0)
                if pg not in page_map:
                    page_map[pg] = []
                page_map[pg].append(item)

            for pg_idx in sorted(page_map.keys()):
                pg_text_parts = []
                pg_tables = []
                pg_images = []
                pg_formulas = []
                for item in page_map[pg_idx]:
                    itype = item.get("type", "")
                    if itype == "text":
                        pg_text_parts.append(item.get("text", ""))
                    elif itype == "table":
                        pg_tables.append({"index": len(pg_tables), "html": item.get("html", ""), "markdown": item.get("text", "")})
                    elif itype == "image":
                        pg_images.append({"index": len(pg_images), "alt_text": item.get("alt_text", "")})
                    elif itype == "formula":
                        pg_formulas.append({"index": len(pg_formulas), "latex": item.get("text", "")})

                result.pages.append({
                    "page_number": pg_idx + 1,
                    "text": "\n".join(pg_text_parts),
                    "tables": pg_tables,
                    "images": pg_images,
                    "formulas": pg_formulas
                })

        # 填充结果
        result.success = True
        result.full_text = full_text.strip()
        result.markdown = md_content.strip() if md_content else ""
        result.images = images
        result.tables = tables
        result.formulas = formulas
        result.metadata = {
            "page_count": ds.page_count if hasattr(ds, 'page_count') else len(result.pages),
            "file_size": len(pdf_bytes),
            "parse_time_ms": int((time.time() - start_time) * 1000),
            "engine": "mineru",
            "version": MINERU_VERSION
        }

        logger.info(f"解析完成 - 耗时: {result.metadata['parse_time_ms']}ms, "
                     f"文本: {len(result.full_text)}字符, "
                     f"图片: {len(images)}, 表格: {len(tables)}, 公式: {len(formulas)}")

    except Exception as e:
        result.success = False
        result.error = f"MinerU 解析异常: {str(e)}\n{traceback.format_exc()}"
        logger.error(result.error)

    finally:
        # 清理临时目录
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except Exception as e:
                logger.warning(f"清理临时目录失败: {e}")

    return result


def parse_pdf_fallback(pdf_bytes: bytes) -> ParseResult:
    """
    降级方案：当 MinerU 不可用时，使用 PyMuPDF (fitz) 做基础提取
    如果 PyMuPDF 也不可用，返回失败
    """
    result = ParseResult()
    start_time = time.time()

    try:
        import fitz  # PyMuPDF
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")

        text_parts = []
        for page_idx in range(len(doc)):
            page = doc[page_idx]
            text = page.get_text()
            text_parts.append(text)

            result.pages.append({
                "page_number": page_idx + 1,
                "text": text.strip(),
                "tables": [],
                "images": [],
                "formulas": []
            })

        result.success = True
        result.full_text = "\n\n".join(text_parts).strip()
        result.metadata = {
            "page_count": len(doc),
            "file_size": len(pdf_bytes),
            "parse_time_ms": int((time.time() - start_time) * 1000),
            "engine": "pymupdf_fallback"
        }
        doc.close()

    except ImportError:
        result.success = False
        result.error = "MinerU 和 PyMuPDF 均不可用"
    except Exception as e:
        result.success = False
        result.error = f"降级解析失败: {str(e)}"

    return result


# ==================== API 端点 ====================

@app.get("/health")
async def health_check():
    """健康检查 + MinerU 可用性"""
    return {
        "status": "ok",
        "mineru_available": MINERU_AVAILABLE,
        "mineru_version": MINERU_VERSION,
        "mineru_error": MINERU_ERROR if not MINERU_AVAILABLE else ""
    }


@app.get("/check-install")
async def check_install():
    """检查 MinerU 安装状态"""
    checks = {
        "mineru_installed": MINERU_AVAILABLE,
        "mineru_version": MINERU_VERSION,
        "fastapi_installed": True,
        "uvicorn_installed": True,
    }

    # 检查模型文件
    model_dirs = []
    home = Path.home()
    candidates = [
        home / ".magic-pdf" / "models",
        home / ".cache" / "magic-pdf" / "models",
    ]
    for cd in candidates:
        if cd.exists():
            model_dirs.append(str(cd))

    checks["model_dirs"] = model_dirs
    checks["models_downloaded"] = len(model_dirs) > 0

    return checks


@app.post("/parse")
async def parse_document(
    file: UploadFile = File(...),
    extract_images: bool = Form(default=True),
    fallback: bool = Form(default=True)
):
    """
    解析 PDF 文档
    - 优先使用 MinerU 进行完整解析（文本+图片+表格+公式）
    - MinerU 不可用时降级到 PyMuPDF（仅文本）
    """
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="仅支持 PDF 文件")

    try:
        pdf_bytes = await file.read()
        logger.info(f"收到解析请求: {file.filename}, 大小: {len(pdf_bytes)} 字节")
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"读取文件失败: {str(e)}")

    # 优先 MinerU
    if MINERU_AVAILABLE:
        result = parse_pdf_mineru(pdf_bytes, extract_images=extract_images)
        if result.success:
            return _result_to_json(result)
        else:
            logger.warning(f"MinerU 解析失败，尝试降级: {result.error}")

    # 降级
    if fallback:
        result = parse_pdf_fallback(pdf_bytes)
        return _result_to_json(result)

    # 不降级，返回 MinerU 错误
    return JSONResponse(content={
        "success": False,
        "error": MINERU_ERROR if not MINERU_AVAILABLE else "MinerU 解析失败",
        "full_text": "",
        "markdown": "",
        "images": [],
        "tables": [],
        "formulas": [],
        "pages": [],
        "metadata": {}
    })


@app.post("/parse-base64")
async def parse_document_base64(payload: dict):
    """
    通过 Base64 解析 PDF（兼容 VB.NET WebClient 调用方式）
    payload: { "base64": "...", "extract_images": true, "fallback": true }
    """
    b64_data = payload.get("base64", "")
    extract_images = payload.get("extract_images", True)
    fallback = payload.get("fallback", True)

    if not b64_data:
        raise HTTPException(status_code=400, detail="缺少 base64 数据")

    try:
        pdf_bytes = base64.b64decode(b64_data)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Base64 解码失败: {str(e)}")

    if MINERU_AVAILABLE:
        result = parse_pdf_mineru(pdf_bytes, extract_images=extract_images)
        if result.success:
            return _result_to_json(result)

    if fallback:
        result = parse_pdf_fallback(pdf_bytes)
        return _result_to_json(result)

    return JSONResponse(content={
        "success": False,
        "error": MINERU_ERROR if not MINERU_AVAILABLE else "MinerU 解析失败",
        "full_text": "",
        "markdown": "",
        "images": [],
        "tables": [],
        "formulas": [],
        "pages": [],
        "metadata": {}
    })


@app.post("/install")
async def install_mineru():
    """
    安装 MinerU（通过 pip）
    返回安装进度和结果
    """
    import subprocess

    try:
        logger.info("开始安装 MinerU...")
        proc = subprocess.run(
            [sys.executable, "-m", "pip", "install", "mineru[core]", "--quiet"],
            capture_output=True,
            text=True,
            timeout=600  # 10 分钟超时
        )

        if proc.returncode == 0:
            logger.info("MinerU 安装成功")
            return {"success": True, "message": "MinerU 安装成功，请重启服务"}
        else:
            return {"success": False, "error": proc.stderr}

    except subprocess.TimeoutExpired:
        return {"success": False, "error": "安装超时（超过10分钟）"}
    except Exception as e:
        return {"success": False, "error": str(e)}


# ==================== 辅助函数 ====================

def _result_to_json(result: ParseResult) -> dict:
    """将 ParseResult 转为 JSON 可序列化字典"""
    return {
        "success": result.success,
        "full_text": result.full_text,
        "markdown": result.markdown,
        "images": result.images,
        "tables": result.tables,
        "formulas": result.formulas,
        "pages": result.pages,
        "metadata": result.metadata,
        "error": result.error
    }


# ==================== 启动入口 ====================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="MinerU Document Parser HTTP Service")
    parser.add_argument("--host", default="127.0.0.1", help="监听地址 (默认: 127.0.0.1)")
    parser.add_argument("--port", type=int, default=8036, help="监听端口 (默认: 8036)")
    args = parser.parse_args()

    logger.info(f"MinerU 可用: {MINERU_AVAILABLE}, 版本: {MINERU_VERSION}")
    if not MINERU_AVAILABLE:
        logger.warning(f"MinerU 未安装: {MINERU_ERROR}")
        logger.warning("降级模式: 仅支持 PyMuPDF 基础文本提取")

    logger.info(f"启动服务: http://{args.host}:{args.port}")
    uvicorn.run(app, host=args.host, port=args.port, log_level="info")
