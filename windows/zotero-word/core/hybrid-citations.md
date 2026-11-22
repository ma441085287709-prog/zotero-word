# 批量字段与原生验收

日常任务直接走 [Windows 运行入口](run-flow.md)。文件工具仍可分别使用。

- `hybrid-inspect` 给出活动正文、表格中的精确坐标和已有字段。
- `hybrid-plan` 接受已核对的组、真实库快照及可选目标样式。原有动态域按身份保留。非支持故事区及跨公式、修订、受保护域的范围按位置报告。
- `hybrid-apply` 只创建新的独立 DOCX。批量写入工作副本并检查计划内外的文字、修订、对象和域。状态为 `PENDING_NATIVE_VALIDATION`。
- `hybrid-audit` 核对磁盘文件的保护边界。正式入口在同一副本上执行 WinWord 刷新、读回活动域和书目，再由 Word 保存并关闭。
