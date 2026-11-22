# macOS 原生执行

Codex 在本机桌面会话调用[正式入口](../core/run-flow.md)。Word 和 Zotero 需要安装在这台 Mac，Zotero 需要可用的 MacWord 集成和目标 CSL 样式。日常连接使用 macOS 随附的 `osascript`。授权由系统管理，拒绝时报告具体应用和报错。

1. 工作副本写好后，Word 按完整路径打开它。若已打开，Word 激活现有窗口并保留其中的编辑。后续操作只针对这份副本。Zotero 选中第一处引文时，脚本仅按副本文件名排除同名歧义，再核对选中的引文。[Word AppleScript](../scripts/word.applescript) · [MacWord 文档绑定](https://github.com/zotero/zotero-word-for-mac-integration/blob/main/build/src/application.m)
2. 读取个人文库时，已有 Zotero 原生会话就直接复用。首次读取优先走本机 API，无法取得完整文库时自动转回原生读库。首项原生任务通过「工具 → 开发者 → Run JavaScript」打开本次专用编辑器，加载包内代码。粘贴后恢复剪贴板，用户同时复制的新内容优先保留。获准导入与刷新共用这个执行器。[读库连接](../core/scripts/mac_library_read.py) · [桌面连接代码](../core/scripts/mac_connection.py) · [Zotero 会话代码](../scripts/zotero-session.js)
3. [MacWord 脚本](../scripts/zotero-native.js) 获得当前文档句柄并选中一条已知引文。Word 单独读回同一文件中选中字段的代码。身份一致后，Zotero 才切换必要的样式偏好、按需添加动态书目并刷新。
4. Zotero 写入刷新回执。Word 对同一工作副本确认保存并关闭。执行器记录该文件的指纹和 `word_window=closed`，用这份文件交付。

中断后在同一目录执行 `run-next`。入口复用当前任务的刷新回执，在已绑定的 Word 窗口中原位保存。已有保存回执时只续办关闭，关闭已确认时直接交付。[Word 执行器](../core/scripts/mac_desktop.py)

可直接从解压目录运行。macOS 首次询问桌面控制权限时，按系统显示的应用名称授权。[本地连接说明](../core/local-connections.md)
