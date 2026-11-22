# zotero-word 2.4.4 windows 使用说明

用于同一 Windows 交互桌面上的 Microsoft Word 桌面版、Zotero 桌面版，以及能执行本机命令的 Codex 或 DeepSeek Harness。使用 Zotero 随附的 Word 集成和目标期刊 CSL 样式。运行环境需要 Windows PowerShell 5.1、Windows Python 3.10 或更高版本及该解释器中的 `lxml`。

## 放入独立项目

1. 将 ZIP 解压到本地目录。若 Windows 在下载文件属性中显示“解除锁定”，先核对来源和提供方给出的 SHA-256，按本机软件策略处理该标记后再解压。企业要求签名时请使用组织批准的分发方式。
2. 在解压出的 `zotero-word` 文件夹中打开 Windows PowerShell。将下列路径换成自己项目的技能目录后执行。安装器创建 `zotero-word` 子目录，遇到已有版本会停止。DeepSeek Harness 使用其已配置的技能目录。

   ```powershell
   .\Install-Windows.ps1 -SkillsRoot "D:\Project\.agents\skills"
   ```

3. 从助手发起实际稿件任务即可。Word 与 Zotero 使用同一登录账户和桌面会话。需要 Code Execution 权限时沿用本机宿主的正常授权。连接代码已在包内，编辑器输入和运行由连接器完成。

## 给助手的任务示例

> 使用 zotero-word，将这份 Word 稿的引文和 References 按 Chemical Engineering Journal（CEJ）样式整理。保留原稿，Zotero 刷新后用 Word 保存并关闭工作副本，把这份文件给我。

如需允许添加确有 DOI 的缺失文献，可在请求中明确授权补库。开始前可保留已打开的普通 Word 文档。如果自己正在 Zotero 的“运行 JavaScript”窗口里写代码，请先保存其中的工作并关闭该窗口。程序会使用自己新开的窗口。

## 取回结果或续办

- `COMPLETE` 后取回执中 `document` 指向的 DOCX。回执给出引文数、书目条数、Word 窗口状态和实际耗时。
- `NEEDS_MATCHING` 时按列出的位置确认待匹配文献。若出现 `HOLD`，保留工作目录，由助手从该目录执行 `run-next`。已有刷新回执时只补 Word 保存或关闭。刷新已发出且结果不明时保留现场，避免重复刷新。
- 若 Word 在打开副本时返回 `WORD_CONTROL_DENIED`，在同一 Windows 登录桌面使用获准控制 Word 的宿主或 PowerShell，执行回执中完整的 `word_control_error.resume_command`。命令已包含工作目录和状态文件路径。
- 用户另行请求核对页面或书目字号时，对已完成的工作目录执行 `run-pages`。[命令与回执](core/run-flow.md)

本包使用 Windows Word COM 和 Zotero 自带的 WinWord 集成。本次打包只核对了文件内容。
