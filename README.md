# Codex PDF Trans PPT

把可复制文字的 PDF 转成可编辑 PPT。

当前版本：v0.2.0

## 文件

- `mypaper.pdf`：原始 PDF
- `mypaper_clean_no_answers.pptx`：当前最终生成的 PPT
- `pdf_text_to_ppt.py`：转换脚本
- `web_app.py`：本地可视化页面
- `templates/`：页面模板
- `static/`：页面样式和交互
- `requirements.txt`：Python 依赖

## Web 页面

```bash
python -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python web_app.py
```

打开：

```text
http://127.0.0.1:5000
```

页面支持上传 PDF、选择跳过答案页策略、选择清爽排版或贴近原文位置排版，并下载生成的 PPTX。

## 命令行

```bash
.venv/bin/python pdf_text_to_ppt.py mypaper.pdf mypaper_clean_no_answers.pptx --skip-from-page 29 --layout clean
```

## 当前策略

- 输出为 16:9 PPT。
- 跳过 PDF 第 29 页及后续答案页。
- 正文合并 PDF 物理换行，减少无意义换行。
- 标题、编号、题目保留必要换行。
- PDF 表格/图表文字用占位符替代。
- PDF 图片尽量缩小放入相关页。
