---
name: gongwen-format
description: Format and check Chinese official-document Word files against configurable 公文格式 requirements. Use when a user asks to standardize, 核对, 整理, 排版, or convert a .docx/.doc into 标准公文格式, especially with GB/T 9704-2012-style page setup, 标题/正文/附件/落款/日期/页码, or a format review report.
---

# 公文格式智能核对与规范化

Use this skill when the user provides or references a `.docx` or legacy `.doc` file and asks to整理成标准公文格式、核对公文格式、按公文要求排版、生成核对报告, or similar.

## Workflow

1. **规则库**  
   Use `references/gongwen_rules.json` as the local rule set. It encodes page setup, fonts, paragraph spacing, semantic patterns, checks, and manual-review items.

2. **Word 结构解析**  
   Use the bundled script to inspect Word paragraphs, runs, sections, margins, and document structure. Do not rely only on copied visible text.

3. **语义识别**  
   Identify title, main receiver, body, heading levels, attachment note, issuer, and date using the semantic patterns in the rule library.

4. **自动排版 + 核对报告**  
   For `.doc`, first convert it to `.docx` with LibreOffice/soffice or macOS `textutil`. Apply deterministic formatting to the `.docx`, then produce both:
   - a standardized Word document named `原文件名-标准公文格式.docx`
   - a review report named `原文件名-格式核对报告.docx`

## Command

Run from the skill directory or a workspace containing this skill:

```bash
python gongwen-format/scripts/format_gongwen_docx.py path/to/input.docx -o output
python gongwen-format/scripts/format_gongwen_docx.py path/to/input.doc -o output
```

Install dependencies first when needed:

```bash
pip install -r gongwen-format/requirements.txt
```

For `.doc` conversion on Linux servers, install LibreOffice/soffice. On macOS, system `textutil` can be used.

## File Tool Mode

This skill can be wrapped by a file-upload/file-download agent tool:

```bash
python gongwen-format/scripts/file_channel_cli.py --input {input_file} --output-dir {output_dir}
```

The tool should pass the uploaded `.docx` or `.doc` path as `{input_file}` and return the files listed in the JSON output as attachments. A generic channel manifest is available at `references/file_channel_tool_manifest.json`.

## Font Notes

The default rules reference common Chinese official-document fonts such as 方正小标宋、仿宋_GB2312、楷体_GB2312. Font files are not bundled in this public package because font redistribution may require separate authorization.

If exact font rendering is required, install legally licensed fonts on the host system, or place locally authorized font files in `assets/fonts` with the names expected by `scripts/install_fonts.sh`.

## Behavior

The script applies these high-confidence corrections:

- A4 page setup with configurable margins.
- Title: configured title font, 2号/22pt, centered, not bold.
- Body: configured body font, 3号/16pt, fixed 28pt line spacing, first-line indent of 2 Chinese characters.
- First-level headings: configured level-1 heading font.
- Second-level headings: configured level-2 heading font.
- Third/fourth-level headings: configured body-style heading font.
- Attachment note, issuer, date, and page number formatting.

The report flags manual-review items instead of silently changing meaning-sensitive content:

- 文种是否规范, including 请示/报告/函混用.
- 标题是否存在介词或文种重叠.
- 主送机关是否符合行文关系.
- 附件说明和附件正文是否一致.
- 保密属性、内容口径、审批责任和实体盖章效果.

## Important Boundaries

- Keep original substantive text unchanged unless the user explicitly asks for rewriting.
- Do not process classified, confidential, or sensitive documents through untrusted services.
- Treat the output as pre-submission formatting assistance; final公文审核仍由起草部门、归口部门和审批链条负责.
