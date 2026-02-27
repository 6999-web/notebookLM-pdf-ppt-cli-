# PPT-Edit-CLI

这是一个基于 AI 的 PPT 还原工具，可将图片版 PPT/PDF 转换为可编辑的 .pptx 文件。

## 核心能力
- **OpenCV 视觉预处理**：背景提取与元素分割。
- **PaddleOCR 识别**：精准提取文字坐标与内容。
- **Gemini 语义重构**：利用大模型纠错、归组并推断样式（字号、对齐、颜色）。
- **图片层级生成**：保持背景美观的同时提供可编辑文字层。

## 安装步骤
1. 安装 Python 3.10+
2. 安装依赖：
   ```bash
   pip install -r requirements.txt
   pip install paddlepaddle  # CPU版本
   # 或者 pip install paddlepaddle-gpu # GPU版本
   ```
3. 配置环境变量：
   将你的 Gemini API Key 写入 `.env` 文件：
   ```text
   GEMINI_API_KEY=your_api_key_here
   ```

## 使用方法
```bash
python -m ppt-edit-cli.main -i ./input_images -o result.pptx
```

## 提示词建议 (针对艺术字)
像素材中那些极大的书法字（如“怒”、“壮”），本工具目前采用“背景全图+透明/覆盖文字层”的策略。
如果你希望更完美的视觉还原，建议在 LLM 识别到 `is_artistic: true` 时，在生成逻辑中保留原始切片。
