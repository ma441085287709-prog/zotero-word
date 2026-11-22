# DeepSeek Harness macOS

在有本地命令执行能力的 Harness 桌面会话，使用[同一个 Python 入口](../core/run-flow.md)和[MacWord 执行器](codex-macos.md)。Python、Word 和 Zotero 必须运行在同一台 Mac 的同一个登录用户会话。授予 Harness 的系统权限以当前宿主实际获得的权限为准。

读 `--brief` 的状态、待匹配位置或原生错误。提交结果时使用本次 `run.json` 记录的文件和数据。DeepSeek Harness 的独立宿主验证留待客户试用。

当 `word_control_error.resume_command` 出现时，引文已写进 `word_control_error.working_copy`，刷新尚未发出。将这份副本的路径、`next_action` 和完整命令一次性告诉用户，当前任务保留为 `HOLD`。等待本宿主获准控制 Word，或由用户选定同一台 Mac、同一登录账户下已有权限的宿主，执行该命令一次。它会从原运行目录续办，取得 `COMPLETE` 后交付回执中的 `document`。

如果用户已通过 Word 界面完成刷新、保存并关闭，直接说明该文件为手工完成，保留原运行记录。
