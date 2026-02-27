# PPT 转换与 AI 还原工具使用指南

这篇指南包含两个核心工具：
1. **NotebookLM PPT 转换器**: 处理 NotebookLM 生成的结构化 PPT。
2. **AI PPT 还原器 (ppt-edit-cli)**: 处理图片版 PPT 或 PDF，通过 OCR 和 AI 还原为可编辑文字。

---

## �️ 工具 1：NotebookLM PPT 转换器
用于把 NotebookLM 生成的 PPTX 转换成可以自由编辑排版的格式。

### 日常操作步骤
1. **移动文件**: 将导出的 `.pptx` 放到 `c:\Users\xxzx-admin\Desktop\实用工具`。
2. **运行转换**:
   ```powershell
   cd c:\Users\xxzx-admin\Desktop\实用工具
   .\.venv\Scripts\python.exe notebook_editor\main.py "原始文件.pptx"
   ```
3. **输出**: 生成 `原始文件_editable.pptx`。

---

## 🤖 工具 2：AI PPT 还原器 (ppt-edit-cli)
用于将不可编辑的**图片、截图或 PDF 转图片**还原为带文字层、背景图的 PPT。

### 准备工作
此工具依赖大模型识别，建议在 `.env` 文件中配置 `GEMINI_API_KEY` 以获得最佳的排版还原效果。

### 运行环境安装 (仅需一次)
如果在当前电脑尚未安装关键依赖，请运行：
```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

### 日常转换步骤
1. **准备文件**: 将 PPT 每一页的截图放到一个文件夹中，或者**直接准备 PDF 文件**。
2. **执行还原**:
   ```powershell
   cd c:\Users\xxzx-admin\Desktop\实用工具
   # 处理文件夹
   .\.venv\Scripts\python.exe -m ppt_edit_cli.main -i ./input_images -o "还原后的PPT.pptx"
   # 直接处理 PDF
   .\.venv\Scripts\python.exe -m ppt_edit_cli.main -i "./我的文档.pdf" -o "还原后的PPT.pptx"
   ```
3. **参数说明**:
   - `-i`: 输入文件夹路径或 **PDF 文件路径**。
   - `-o`: 输出 PPT 文件名。
   - `-k`: (可选) Gemini API Key。

### 特别：高级还原特性
- **智慧去底**: 自动擦除原图中的文字内容（Inpainting），确保新生成的文字层下方没有模糊的“残影”。
- **水印移除**: 针对 NotebookLM 自动识别并移除右下角水印。
- **元素识别**: 能够自动区分标题、正文、列表、甚至识别图片区域。
- **广泛兼容**: 支持 PNG/JPG 文件夹或单体 PDF 文件。

---

## 💡 常见问题与建议
- **文件名空格**: 如果文件名有空格，请务必用英文双引号包裹。
- **环境隔离**: 所有工具均运行在当前文件夹的 `.venv` 虚拟环境中，不会污染系统环境。
- **性能提示**: AI 还原工具处理大量图片时可能需要几十秒到几分钟，请耐心等待进度条完成。
