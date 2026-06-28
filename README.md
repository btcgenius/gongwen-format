# gongwen-format skill

一个用于中文公文 Word 文件格式核对与规范化的 Codex Skill。它通过本地规则库解析 `.docx` / `.doc` 文档结构，自动应用页面、字体、段落、附件、落款、日期、页码等格式规则，并生成格式核对报告。

## 包含内容

```text
gongwen-format/
├── SKILL.md
├── agents/openai.yaml
├── references/
│   ├── gongwen_rules.json
│   └── file_channel_tool_manifest.json
├── scripts/
│   ├── format_gongwen_docx.py
│   ├── file_channel_cli.py
│   ├── check_fonts.py
│   └── install_fonts.sh
├── assets/fonts/
└── requirements.txt
```

## 快速使用

安装依赖：

```bash
pip install -r gongwen-format/requirements.txt
```

处理 `.docx` 文件：

```bash
python gongwen-format/scripts/format_gongwen_docx.py path/to/input.docx -o output
```

处理文件上传/下载通道：

```bash
python gongwen-format/scripts/file_channel_cli.py --input path/to/input.docx --output-dir output
```

## 字体说明

本公开包不包含字体文件。默认规则引用方正小标宋、仿宋_GB2312、楷体_GB2312 等常见中文公文字体；如果需要精确渲染，请在目标机器安装合法授权字体，或将字体文件自行放入 `gongwen-format/assets/fonts/` 后运行：

```bash
bash gongwen-format/scripts/install_fonts.sh
```

## 安全边界

- 不要把涉密、敏感或未经授权的内部文件提交到公开仓库。
- 规则库中的格式要求可按本单位制度自行调整。
- 工具只做格式辅助处理，不替代内容审核、审批责任、行文关系判断或最终用印确认。
