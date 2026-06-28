#!/usr/bin/env python3
"""Native file-channel entrypoint for the self-contained gongwen-format skill."""

from __future__ import annotations

import argparse
import json
import shutil
import sys
import uuid
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
RULES_PATH = SKILL_DIR / "references" / "gongwen_rules.json"

sys.path.insert(0, str(SCRIPT_DIR))
from format_gongwen_docx import Finding, format_document, prepare_input_document, write_report  # noqa: E402


def make_zip(zip_path: Path, files: list[Path]) -> None:
    with ZipFile(zip_path, "w", ZIP_DEFLATED) as zf:
        for file in files:
            zf.write(file, arcname=file.name)


def main() -> None:
    parser = argparse.ArgumentParser(description="公文格式智能核对与规范化文件通道命令工具")
    parser.add_argument("--input", required=True, type=Path, help="用户上传的 .docx 或 .doc 文件")
    parser.add_argument("--output-dir", required=True, type=Path, help="文件通道输出目录")
    parser.add_argument("--job-id", default=None, help="可选任务 ID")
    args = parser.parse_args()

    if not args.input.exists() or args.input.suffix.lower() not in {".docx", ".doc"}:
        raise SystemExit("input must be an existing .docx or .doc file")

    job_id = args.job_id or str(uuid.uuid4())
    job_dir = args.output_dir / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    working_input = job_dir / args.input.name
    shutil.copy2(args.input, working_input)

    standard_docx = job_dir / f"{working_input.stem}-标准公文格式.docx"
    report_docx = job_dir / f"{working_input.stem}-格式核对报告.docx"
    prepared_input, converted = prepare_input_document(working_input, job_dir)
    ctx = format_document(prepared_input, standard_docx, report_docx, RULES_PATH)

    if converted:
        ctx.findings.insert(0, Finding("格式转换", None, working_input.name, "已先将 .doc 转换为 .docx 后再进行公文格式处理"))
        write_report(report_docx, working_input, standard_docx, ctx)

    zip_path = job_dir / "公文格式处理结果.zip"
    make_zip(zip_path, [standard_docx, report_docx])

    print(json.dumps({
        "status": "ok",
        "job_id": job_id,
        "recognized_paragraphs": len(ctx.roles),
        "finding_count": len(ctx.findings),
        "files": {
            "standard_docx": str(standard_docx),
            "review_report": str(report_docx),
            "zip": str(zip_path),
        },
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
