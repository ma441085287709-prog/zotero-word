---
name: zotero-word
description: 在 Windows 的 Word 和 Zotero 中按期刊样式整理动态引文和 References，按 DOI、链接或作者年份匹配文献，由 Zotero 刷新，Word 保存、关闭后直接交付工作副本。适用于可调用 Windows 桌面的 Codex 或 DeepSeek Harness。
metadata:
  version: "2.4.4"
  build: "direct-save.2-handoff.1-dispatch-fixed"
  release_channel: customer-release
  platform: windows
  test_status: NOT_RUN_AT_USER_REQUEST
---

# zotero-word 2.4.4 windows

客户安装和首次使用见[使用说明](INSTALL-WINDOWS.md)。用户要求处理稿件时，在本技能目录调用正式入口。同一 Windows 桌面上的 Codex 和 DeepSeek Harness 使用同一命令。

```powershell
py -3 -X utf8 .\core\scripts\zotero_docx.py run-start "D:\Paper\稿件.docx" --journal "CEJ" --run "D:\Paper\引文工作目录" --output "D:\Paper\引文状态.json" --brief
```

入口保留原稿和只读快照，在同一份工作副本中写入动态域。Zotero 的 WinWord 集成应用期刊样式、刷新并读取真实引文和书目。Word 保存并关闭这份副本，回执中的 `document` 即交付稿。CEJ 和 ES&T 可直接使用期刊简称，其他已安装样式可用 `--style-id` 指定。

- ScienceDirect 的 PII 链接先按出版商标识查 DOI，再匹配已有文库。化学电荷和引文之间的原有空格保留在字段外。
- 需要补库且用户已授权时，加 `--import-missing`。需要人工确认的位置写入 `matching.json`。用同一目录的 `run-next --resolve 位置编号=真实条目URI` 继续，同一组的多篇条目按引用顺序传入。[运行说明](core/run-flow.md)
- `COMPLETE` 后交付 `document`，报告引文数、书目条数和实际耗时。常规任务到此结束。
- `HOLD` 时报告失败阶段、`reason_code` 和可用的 `diagnostic`。沿用原目录执行 `run-next`。若刷新已经确认，入口只续办 Word 尚未完成的保存或关闭。若刷新已发出、结果未知，保留这份副本及现场。
- 若 Word 在打开副本时拒绝宿主访问，且回执包含 `word_control_error.resume_command`，向用户提供工作副本路径及完整的 PowerShell 续办命令。[Windows 续办说明](adapters/deepseek-harness-windows.md)
- 用户另行要求核对页面时使用 `run-pages`。另行委托全文排版或清洁稿时才创建派生副本并核验。[可选全文格式](core/optional-layout.md)
