---
name: zotero-word
description: 在 macOS 的 Word 和 Zotero 中按期刊样式整理动态引文和 References，按 DOI、链接或作者年份匹配文献，由 Zotero 刷新，Word 保存、关闭后直接交付工作副本。
metadata:
  version: "2.4.4"
  build: "direct-save.2-handoff.1"
  release_channel: customer-release
  platform: macos
  test_status: NOT_RUN_AT_USER_REQUEST
---

# zotero-word 2.4.4  mac

安装和首次使用见[客户使用说明](使用说明.md)。

用户要求处理稿件时，在解压后的 `zotero-word` 目录使用正式入口。Codex 和有本机执行能力的 DeepSeek Harness 共用该入口。它会选用本机已有、可加载 lxml 的 Python 3.10+。

```sh
python3 core/scripts/zotero_docx.py run-start "/路径/稿件.docx" --journal "ES&T" --run "/路径/引用工作目录" --output "/路径/引文状态.json" --brief
```

入口保留只读原稿快照，在唯一的工作副本上批量写入引文。Zotero 的 MacWord 集成刷新该文件，Word 确认保存并关闭后即交付。完成回执中的 `document` 就是这份工作副本。新字段支持正文和可编辑表格内能准确定位的范围。其它期刊可传 `--style-id` 指定已安装的 CSL 样式。

- 遇到 ScienceDirect PII 链接，默认先匹配已有文库。需要新增条目且用户已授权时加 `--import-missing`。待定位置及上下文写入 `matching.json`。让用户确认实际条目后，使用同一目录的 `run-next --resolve 位置编号=条目URI` 继续。[运行说明](core/run-flow.md)
- 写入动态域时保留化学电荷与引文之间的原有空格，清理转换链接留下的空格格式。已完成稿的书目字号在文件中无法解析时，书目格式检查可读取 Word 的实际字号。[引文与书目格式](core/manuscript-format.md)
- `COMPLETE` 表示已取得 Zotero 刷新回执，以及同一工作副本的 Word 保存、关闭回执。向用户交付该文件，并报告引文数、书目条数和实际耗时。
- `HOLD` 时报告 `resume_from`、`reason_code` 和可用的 `diagnostic`。Word 拒绝自动化控制时，查看 `word_control_error` 的操作和 Apple Event 错误码。保留运行目录，以便续办。DeepSeek 收到 `word_control_error.resume_command` 时，直接给出工作副本路径、所需 Word 权限和这条命令。[DeepSeek 续办说明](adapters/deepseek-harness-macos.md)
- 在原运行目录执行 `run-next`。有刷新回执就从 Word 保存或关闭继续，两份 Word 回执齐全时直接交付。刷新命令结果未知时保留现场。[续办命令](core/run-flow.md)

原稿可保持打开。普通任务在交付引文稿后结束。用户另行委托排版时再创建排版副本。`layout-audit` 会对实际排版稿执行原生重开核验。对另行制作的清洁稿，可用 `run-verify-copy` 加 `--clean` 生成它自己的回执。运行目录存放副本及恢复记录，执行代码随包提供。[Mac 原生执行说明](adapters/codex-macos.md)
