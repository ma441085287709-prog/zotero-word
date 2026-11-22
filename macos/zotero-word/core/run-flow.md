# macOS 正式入口

`run-start 原稿 --run 目录 --journal 期刊 --output 状态.json` 建立或继续同一任务。运行目录用于原稿快照、定位资料、工作副本和原生执行记录。状态 JSON 放在运行目录外或其顶层，单独命名。已有旧版运行可用 `run-status` 查阅，保留其原始记录。

`run-next --run 目录 --output 状态.json` 复用同一任务。可用 `--style-id` 指定已安装样式，用 `--library` 提供指定文库快照，用 `--resolve` 确认待定位置的真实条目。需要改动文字范围或标记非引文时，仍可用 `--entries` 提交复核记录。默认读取本机个人文库。按任务授权补库时加 `--import-missing`，可在续办时用 `--no-import-missing` 关闭。只有匹配当前个人文库的任务会写库。对尚未匹配的 ScienceDirect 链接，入口先按出版商 PII 查文库内的映射，再核对 Crossref 的精确标识，必要时核对出版商文章页的标识和 DOI。查实的映射缓存到当前任务，同一 PII 在本批次只查询一次。找到已有 DOI 即复用条目。只有授权补库后才查询缺失文献的完整元数据并入库。遇到歧义、证书或数据源错误时，可在 `matching.json` 中查看位置及具体原因。`--reload-library` 重新读取本机文库。程序复用已匹配的位置，并按书目信息召回 PII 候选，以完全相同的 PII 和唯一 DOI 确认。

例如，从 `matching.json` 取出待定位置编号和所选文库条目的实际 URI，替换下列示例值。同一位置需要多篇时，按引用顺序重复 `--resolve`。

```sh
python3 core/scripts/zotero_docx.py run-next --run "/路径/引用工作目录" --resolve "位置编号=http://zotero.org/users/实际用户ID/items/实际条目KEY" --output "/路径/引文状态.json" --brief
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

同一运行由文件锁防止并发写入。原稿只读，`source-snapshots/original.docx` 是后台保护快照。唯一交给 Word 与 Zotero 的稿件是 `citations-<标识>.docx`。Zotero 在该文件中更新字段与书目，Word 确认保存并关闭。此时 `run.json` 记录 `COMPLETE`，`document` 和 `paths.delivery_document` 均指向该工作副本。

中断后执行同一目录的 `run-next`。已绑定的刷新任务把结果写入 `refreshed.json`。窗口仍有待保存的刷新结果时，入口调用 Word 原位保存。Word 已确认保存时只续办关闭。`saved.json`、`closed.json` 的路径与指纹均对应当前工作副本时，直接完成。若上次保存的答复丢失，入口可用落盘字段与该次刷新回执核对。命令结果仍未知或工作文件与回执不符时，保留现场并返回 `HOLD`。

正常续办只需 `--run` 和 `--output`。兼容参数 `--saved-copy` 可以指明同一运行内已确认刷新的工作文件。另行保存的稿件不参与该运行。直接交付回执中 `native_refresh_observed=true`、`word_window=closed`、`sample_verification=not_run`，引文与书目条数来自本次刷新。旧字段 `native_validation_verified` 记录重开核验状态，直接交付时为 `false`。

另行制作的排版稿或清洁稿使用自己的文件路径。以下命令核对该副本相对已验引文稿的科学正文、受保护对象和动态域，并实际重开。`--clean` 另外检查剩余批注和修订。回执只覆盖该命令核验的内容，全文版式验收使用 `layout-audit`。

```sh
python3 core/scripts/zotero_docx.py run-verify-copy "/路径/清洁稿.docx" --clean --run "/路径/引用工作目录" --output "/路径/清洁稿核验.json" --brief
```

`desktop_counts.zotero_sessions` 记录已收到回执的原生会话数，系统授权弹窗次数无法观察时记为空。各阶段实耗记录在 `stage_seconds` 中。完成回执的 `native_seconds` 记录原生刷新、保存以及从保存到关闭的耗时。`started_at` 至 `finished_at` 为整项任务的历时，包含暂停等待。
