# 3GPP TDocs Reader & Summarizer（兼容入口）

本文件原先保存了一套仅适用于 Windows/PowerShell 的旧版流程。为避免与当前 Linux 实现产生冲突，工作流已统一迁移至 [SKILL.md](SKILL.md)。

当前实现使用：

- Bash、`curl`、`unzip` 管理下载和缓存；
- Python、`openpyxl` 读取会议 Index Excel；
- `lxml`/OOXML 提取 DOCX 正文、表格、VML/WPS 和 Visio 文本；
- LibreOffice 与 `pdftoppm` 作为复杂流程图的渲染回退；
- [scripts/3gpp_tdocs.py](scripts/3gpp_tdocs.py) 执行可重复的数据处理。

请从 [SKILL.md](SKILL.md) 开始，不再执行本文件历史版本中的 PowerShell 命令。
