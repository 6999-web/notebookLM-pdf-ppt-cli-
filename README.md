# 🛠️ PPT 实用工具箱 (Utility Tools)

本项目是一个高效的 PPT 处理工具集合，包含两个核心模块：

## 1. 🤖 AI PPT 还原器 (`ppt_edit_cli`)
**功能**：通过 OCR 和大模型（Gemini）识别图片或 PDF 中的内容，并还原为可直接编辑、排版的 PPT。
- **智慧去底**: 自动擦除原图文字（Inpainting），告别“重影”。
- **水印消除**: 智能移除 NotebookLM 等常见水印。
- **高保真还原**: 自动识别标题、正文、列表、加粗及对齐方式。
- **使用参考**: `.\.venv\Scripts\python.exe -m ppt_edit_cli.main -i <输入PDF或图片文件夹> -o <输出PPT>`

## 2. 📓 NotebookLM 转换器 (`notebook_editor`)
**功能**：专门用于优化 NotebookLM 生成的结构化 PPT。
- **快速转换**: 将不可导出的 PPT 布局转换为标准的可编辑文本框。
- **使用参考**: `.\.venv\Scripts\python.exe notebook_editor\main.py <原始文件.pptx>`

---

## 📂 项目结构
- `ppt_edit_cli/`: AI 还原工具核心代码。
- `notebook_editor/`: NotebookLM 转换工具核心代码。
- `assets/`: 示例文件、原始演示文稿存放处。
- `docs/`: 详细的使用手册。
- `tests/`: 自动化单元测试。
- `requirements.txt`: 全局依赖列表。
- `WORKSPACE` / `BUILD`: Bazel 构建配置文件。

## ⚙️ 快速开始
1. **安装依赖**:
   ```powershell
   .\.venv\Scripts\python.exe -m pip install -r requirements.txt
   ```
2. **详细手册**: 见 [docs/PPT转换工具使用指南.md](docs/PPT转换工具使用指南.md)

---
*Updated: 2026-02-27 | Optimized by Antigravity AI*
