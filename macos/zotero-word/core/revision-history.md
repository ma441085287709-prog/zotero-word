# 修订识别与表格属性历史

读取器和混合写入保护清单使用 `scripts/revision_history.py` 统一分类。`tblGridChange`（旧表格网格）和 `tblPrExChange`（旧行级表格例外属性）都是修订历史；不能只按名称以 `PrChange` 结尾判断。严格预检遇到这些历史仍返回 HOLD，只读原生引文验收可在内存中剔除历史快照查看当前结果；这不会修改实际文件。

用户明确授权在独立副本中接受修订后，可使用 `accept-table-history` 做受限清理。该命令需要非空 `--authorization-evidence`，输入与输出必须不同，输出必须尚不存在。它只删除表格的 `tblGridChange`、`tblPrExChange`、`tblPrChange`、`trPrChange`、`tcPrChange` 历史快照，保留当前属性、网格、文字、公式、非 XML 部件和未来 Track Changes 设置。

这不是通用“接受全部修订”工具。包内任何 XML 部件仍有插入、删除、移动、行/单元格结构修订或其他属性修订时，命令阻断，不发布输出。混合稿件应在已有授权下交由 Word 原生接受修订，再核对完整包中修订标记；不要重复使用外部简化脚本并据此宣称全部清除。

清理回执记录源/输出哈希、修改部件、各类清除数量与授权证据。它只证明受限离线清理，不代替 Word/Zotero 原生刷新与验收。
