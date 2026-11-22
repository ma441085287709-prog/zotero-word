# Windows 正式入口

`run-start 原稿 --run 目录 --journal 期刊 --output 状态.json` 建立或继续同一任务。运行目录用于原稿快照、定位资料、工作副本和原生执行记录。状态 JSON 放在运行目录外或其顶层，单独命名。首次调用示例见 [SKILL.md](../SKILL.md)。已有旧版运行可用 `run-status` 查阅，保留其原始记录。

`run-next --run 目录 --output 状态.json` 复用同一任务。可用 `--style-id` 指定已安装样式，用 `--library` 提供指定文库快照，用 `--resolve` 按位置确认条目。需要改动原文范围或标记非引文时，仍可用 `--entries` 提交完整复核记录。默认读取本机个人文库。按任务授权补库时加 `--import-missing`，可在续办时用 `--no-import-missing` 关闭。只有匹配当前个人文库的任务会写库。遇到 ScienceDirect PII 链接，入口先查文库中的 PII→DOI 映射，再查 Crossref 的精确出版商标识，必要时读取出版商文章页中的标识和 DOI。确认的映射写入当前任务缓存，按 DOI 复用库中条目。只读匹配也走这一路径。遇到歧义、证书或数据源错误时，可在 `matching.json` 中查看位置及具体原因。`--reload-library` 重新读取本机文库。原稿和文库相同时，续办沿用已匹配的位置。按书目信息召回的 Crossref 候选须含完全一致的 PII，并唯一指向一个 DOI。

例如确认一个待定位置，可把下列位置编号及 URI 换成 `matching.json` 中的位置和本机文库的实际条目。同一位置需要多篇时，按引用顺序重复 `--resolve`。

```powershell
py -3 -X utf8 .\core\scripts\zotero_docx.py run-next --run "D:\Paper\引文工作目录" --resolve "cue-id=http://zotero.org/users/local/实际文库ID/items/实际条目KEY" --output "D:\Paper\引文状态.json" --brief
```

位置文件使用 preparation.json 中的源哈希、story、从 1 起算的段落编号和从 0 起算的文字范围。每一组条目按引用顺序列出真实库 URI。非文献链接使用 `non_citations` 将 cue_id 对应到具体理由。原文中的非引文占位提示可在 `non_citation_placeholders` 中用其 cue_sha256 及理由明确排除。未列出的明确 DOI 仍可自动匹配。

书目优先复用已有动态域。需要新建时，可用 `bibliography_paragraph` 指定主文档中的普通空段。自动定位先找现有 References 等书目标题，再找 Disclaimer 或 Publisher’s Note 的起始段，均无时才放在正文末尾。工作副本在选定位置生成一个动态书目域，缺标题时补上 References。已有静态书目须先确定处理范围。Zotero 在设置期刊样式后读取实际域数，Word 再确认目标副本的位置。

```json
{
  "source_sha256": "<preparation.json 中的 document_sha256>",
  "entries": [
    {"id":"citation-1","paragraph":5,"start":11,"end":18,"text":"此处的精确原文","uris":["<真实 Zotero 条目 URI>"]}
  ],
  "non_citations": {}
}
```

同一运行由文件锁防止并发写入。Zotero 对绑定的 `citations-<标识>.docx` 执行一次显式 Refresh。刷新返回后，在同一 WinWord 文档中核对全部引文 ID、每组真实文库条目及顺序、动态书目和引用条目集合。若工作副本原来启用了修订跟踪，在保存前恢复。Word 保存并确认关闭，随后直接交付此文件。原稿和只读快照保留原指纹。

中断后沿用原运行目录。

```powershell
py -3 -X utf8 .\core\scripts\zotero_docx.py run-next --run "D:\Paper\引文工作目录" --output "D:\Paper\引文状态.json" --brief
```

入口读取原生任务的 `refreshed.json`、`saved.json` 和 `closed.json`。刷新已证实时，只完成剩余的 Word 操作。若保存回执缺失但副本已落盘，则比对保存文件中的域、样式和书目与实际刷新结果。已发出刷新而没有最终回执时保留现场。旧版已完成的任务继续指向原来交付的文件。旧版人工另存且缺少本任务刷新证据的候选稿保留在原路径。

当 Word 在写域之后、Zotero 派发之前拒绝打开工作副本时，`word_control_error` 给出工作副本路径、错误和本次任务的完整 PowerShell 续办命令。可在同一登录桌面具备 Word 访问权的宿主继续。

新任务成功时，`completion_scope=zotero_refresh_word_save_close`，`word_window=closed`，`document_sha256` 与 `protection.target_sha256` 绑定这份保存稿。读回引文和书目发生在原生刷新中。常规交付的 `sample_verification=not_run`。旧版已有完整首选项更新回执的任务可续办保存、关闭，以 `zotero_preferences_update_word_save_close` 记录实际更新来源。

另行制作的排版稿或清洁稿使用自己的文件路径。以下命令核对该副本相对已验引文稿的科学正文、受保护对象和动态域，并实际重开。`--clean` 另外检查剩余批注和修订。回执只覆盖该命令核验的内容，全文版式验收使用 `layout-audit`。

```powershell
py -3 -X utf8 .\core\scripts\zotero_docx.py run-verify-copy "D:\Paper\清洁稿.docx" --clean --run "D:\Paper\引用工作目录" --output "D:\Paper\清洁稿核验.json" --brief
```

## 已完成稿的页面复核

用户要求复核版面或需要 Word 的书目实测字号时，使用同一工作目录。

```powershell
py -3 -X utf8 .\core\scripts\zotero_docx.py run-pages --run "D:\Paper\引文工作目录" --output "D:\Paper\页面回执.json" --brief
```

入口比对完成回执和磁盘文件，在独立的只读 Word 会话中逐页生成 PNG，同时读取书目各段的实际字体、字号和段落设置。完整记录位于运行目录的 `visual-qa` 下。相同交付文件可复用已生成且指纹一致的图片。打开图片逐页审阅后再报告版面结果。书目字号在文件内无法解析时，`bibliography-quality` 可采用本次对同一文件、同一段落测得的一致字号。混合字号仍须复核。

`native_update_origin` 记录本任务实际完成的更新。原生刷新回执记录 `scriptVersion` 和 `refreshTiming`。已有目标样式时，`setDocPrefsRequired=false` 且首选项返回时间为空。计时旁路的异常记入 `writeErrors`，续办依据实际刷新、保存和关闭回执。`desktop_counts` 记录桌面进程、原生会话和可观察的首选项提示。`stage_seconds` 记录各阶段实耗，`native_seconds` 单列刷新、保存及保存到关闭的时间。整次历时包含暂停等待。
