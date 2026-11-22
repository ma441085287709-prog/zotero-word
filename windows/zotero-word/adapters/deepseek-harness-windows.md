# DeepSeek Harness on Windows

本地 Harness 使用与 Codex 相同的 [Windows 运行入口](../core/run-flow.md) 和[原生执行器](windows-native.md)。由 Harness 提供其实际 Windows 命令通道和桌面权限。云端网页会话需要能够连接该桌面的宿主。

若 `WORD_CONTROL_DENIED` 发生在绑定工作副本的打开阶段，且刷新尚未派发，回执会给出 `word_control_error.working_copy` 和 `word_control_error.resume_command`。把两项原样交给用户，在能控制该 Word 的同一 Windows 登录桌面上执行 PowerShell 命令。它会读取原运行目录的计划和副本。刷新已确认的任务用同一目录的 `run-next` 续办保存或关闭。
