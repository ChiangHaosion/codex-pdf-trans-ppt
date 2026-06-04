# Codex PDF Trans PPT

把可复制文字的 PDF 转成可编辑 PPT。

当前版本：v0.4.0

## 文件

- `mypaper.pdf`：原始 PDF
- `mypaper_clean_no_answers.pptx`：当前最终生成的 PPT
- `mypaper_answers.pptx`：题目下方附答案的 PPT
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

页面支持上传 PDF、选择转换配置、选择答案展示方式、选择答案页位置、选择清爽排版或贴近原文位置排版，并下载生成的 PPTX。转换任务会在后台执行，任务页会自动刷新状态。

## 命令行

```bash
.venv/bin/python pdf_text_to_ppt.py mypaper.pdf mypaper_clean_no_answers.pptx --skip-from-page 29 --layout clean --profile workbook --answer-mode skip
```

生成题下附答案版本：

```bash
.venv/bin/python pdf_text_to_ppt.py mypaper.pdf mypaper_answers.pptx --skip-from-page 29 --layout clean --profile workbook --answer-mode inline
```

## 测试

```bash
.venv/bin/python -m pytest -q
```

## 当前策略

- 输出为 16:9 PPT。
- 默认从 PDF 第 29 页起跳过（含第 29 页及之后，通常是答案页）。
- 支持 `workbook` 和 `generic` 两种转换配置。
- 支持 `inline` 答案模式：把答案页解析为答案数据源，并插入到对应题目下方。
- `inline` 答案模式下，答案使用红色并带有 Fly In 入口动画。
- 正文合并 PDF 物理换行，减少无意义换行。
- 标题、编号、题目保留必要换行。
- PDF 表格/图表文字用占位符替代。
- PDF 图片会按位置归属到相关 PPT 页；纯图片页也会生成 PPT 页。
- Web 上传支持中文 PDF 文件名，并会把无效 PDF 转换失败记录到任务页。
