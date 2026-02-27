# NotebookLM PPT 转可编辑 PPTX 的 CLI

此工具将 NotebookLM 生成的不易编辑的 PPTX 转换为更易编辑的 PPTX，核心目标是把文本转为可编辑的文本框、把图片重新以可编辑形式放回幻灯片，便于后续在 PowerPoint 中逐步排版。

重要说明
- 不提供去水印的实现。水印的移除通常涉及版权与许可问题，务必确保你对内容拥有合法使用权。若你确实需要无水印版本，请获取授权版本或与原作者/平台联系。
- 该工具的实现偏向“可编辑性优先”的最小可行方案，复杂对象（如图表、表格、SmartArt 等）可能不会被完美还原。

- 安装依赖
- 运行: pip install -r notebook_editor/requirements.txt
- 也可通过 setup.py 安装并获得 notebook-edit CLI：
-   python -m pip install .
- 运行: pip install -r notebook_editor/requirements.txt

使用方法
- 进入项目根目录，执行：
  python notebook_editor/main.py path/to/notebooklm_generated.pptx
- 默认输出为 path/to/notebooklm_generated_editable.pptx
- 指定输出路径：
  python notebook_editor/main.py path/to/notebooklm_generated.pptx -o path/to/output_editable.pptx
- 说明：如果你需要进一步的自定义（如更复杂的布局映射、字体选项、批量处理等），可以告诉我，我再扩展。

实现要点
- 逐幻灯片新建空白布局幻灯片
- 提取文本形状文本，逐个创建可编辑文本框，保持相对位置
- 提取图片，重新插入到新幻灯片的相同位置和尺寸
- 其他对象（如图表、表格等）保持简单实现，以确保核心内容的可编辑性

后续工作建议
- 增强对图表、表格等对象的导出与再现能力
- 增加对字体、字号、颜色的保真映射
- 添加单元测试和对比示例 PPTX
- 如需打包成 npm/CLI 风格，后续可考虑使用 node-pptx 等实现
