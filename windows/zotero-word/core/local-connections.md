# Windows 本机连接

`run-start` 和 `run-next` 通过包内的 Python、PowerShell 与 Zotero 原生执行代码完成整条链路。个人库的首选读通道是 `127.0.0.1:23119/api/`。入口可启动 Zotero 并等待首个库页面；API 不可用时复用同一连接器的原生读库入口。导入仅限已授权、已绑定的个人库与可核实的条目。

Word 通过 COM 控制工作副本。Zotero 通过随附 WinWord 集成刷新和核验。失败时从当前任务的 `run.json` 与 `connection-error.json` 读取阶段、原因和已有回执。恢复和文件交付见[正式流程](run-flow.md)，桌面接口见[Windows 原生执行](../adapters/windows-native.md)。
