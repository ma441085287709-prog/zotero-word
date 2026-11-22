# Windows 桌面执行

一个运行会话复用同一个 STA PowerShell 进程和 Word COM 对象。每次原生写入检查工作文件的完整路径和 Windows 文件标识。解析和匹配完成后，首次需要操作 Word 时才连接。

个人库优先由本机 API 分页读取。冷启动时先等待 Zotero，再决定是否走原生读库。需要运行原生操作时，连接器打开 Zotero 自带的“运行 JavaScript”，仅在新开的窗口或可确认归属及空白状态的窗口输入代码。工具菜单不可见时在 Zotero 主窗口尝试工具菜单助记键。待剪贴板释放，连接器重新核对所有者、任务标记和全文，再从编辑器复制代码逐字核对后执行。后续任务复用会话，检查任务、操作、文档和会话四项回执。[Zotero 的编辑器实现](https://github.com/zotero/zotero/blob/main/chrome/content/zotero/runJS.js)

Zotero 按完整路径请求 WinWord 文档。Word COM 按引文 ID 在对应正文或注释中唯一定位并选中域结果。Zotero 从当前光标独立回读域代码。现有设置已匹配时直接处理书目和刷新。需要改样式时先监听 Zotero，再由同一桌面会话调用绑定 Word 窗口的 Zotero“文档首选项”。工具栏控件不可用时，改用同一 WinWord 集成命令。可以识别的控件自动设置，需人工确认时在已有窗口等一次。首选项和刷新期间仅临时抑制慢更新提示，保留文档原来的自动更新设置。已采用目标样式时直接进入刷新。Zotero 在刷新后读取真实动态域、核对每组条目和书目，Word 保存并关闭当前工作副本。[Zotero 官方 WinWord 集成](https://github.com/zotero/zotero-word-for-windows-integration/blob/main/components/zoteroWinWordIntegration.mjs)

启动刷新任务时，连接器在发送前丢失焦点可尝试唤回 Zotero 一次。Zotero 从包内加载本次版本的脚本，绕过脚本缓存。刷新回执记录脚本版本、样式确认、首选项返回及字段读取时间。未调用首选项时，返回时间为空。计时文件的写入异常记入主回执。

只有另行委托派生稿核验时，才调用 Word 重开和 Zotero Edit Citation 抽查。
