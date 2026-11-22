# 引用、文库与可选排版命令

Windows 引文任务从 [正式运行入口](run-flow.md) 开始。以下是单独请求完整文库核对或全文格式时使用的专项工具。

## 批量库任务

准备阶段可用 `batch-plan "$SOURCE_DOCX" --preparation "$PREPARATION" --library "$LIBRARY" --import-requests "$IMPORT_REQUESTS" --batch-size 20 --output "$BATCH_PLAN" --brief`。不需要导入时省略 `--import-requests`。输入导入请求为数组；具体字段和来源绑定见 [批量规则](batch-operations.md)。确定 DOI 必须相符；出版社 URL 经核实解析为 DOI 时，仍须通过原 URL 字段绑定来源，不能凭一个 `identity_verified` 布尔值替换线索。

保存稿转换后，集合任务另用 `batch-plan "$CITATION_DOCX" --membership-plan "$READY_MEMBERSHIP_PLAN" --library "$LIBRARY" --collection "$COLLECTION" --output "$BATCH_PLAN" --brief`。允许 membership-only 计划；不把转换前后不同哈希混在同一批次。

适配器实际执行后，用 `batch-status "$CURRENT_DOCX" --batch-plan "$BATCH_PLAN" --library "$FRESH_LIBRARY" --collection "$FRESH_COLLECTION" --observations "$ACTUAL_OBSERVATIONS" --output "$BATCH_STATUS" --brief` 重查。没有集合任务或操作观察时省略相应参数。未知写入先重查，明确未写入才允许重新派发。`BATCH_PLANNED` / `BATCH_STATUS` / `OBSERVED_BATCH_EFFECTS` 都不是引用完成；不替代原有 `verify/finalize`。

## A 完成后的选择与可选全文格式

收到用户的全文排版委托后，以已验引文稿及其基准生成绑定信息。

```powershell
& $PYTHON "$SKILL_DIR/core/scripts/zotero_docx.py" layout-offer "$CITATION_DOCX" `
  --baseline "$CITATION_BASELINE" --output "$RUN_DIR/layout-offer.json" --brief
```

`layout-decision` 的 `--decision-source` 填写用户对本稿和期刊的明确请求，仍有效的决定可直接沿用。若用户只要求引文稿，交付后结束。

```powershell
& $PYTHON "$SKILL_DIR/core/scripts/zotero_docx.py" layout-decision "$CITATION_DOCX" `
  --offer "$RUN_DIR/layout-offer.json" --choice yes `
  --decision-source "$ACTUAL_USER_RESPONSE" --output "$RUN_DIR/layout-decision.json" --brief
```

`--choice no` 返回 `DECLINED`，`pending` 仍待答；二者不会读取 B 规则或产出排版稿。命令不能自己推断用户选择。只有有效 yes 后，才运行 `layout-paragraphs "$CITATION_DOCX" --baseline "$CITATION_BASELINE" --decision "$DECISION" --output "$LAYOUT_INVENTORY" --brief`，按实际文字、表格位置与样式审核角色，并核查当前期刊全文格式规则。

B 档案为 `zotero-word.optional-layout-profile.v1`，含 `journal`、`article_type`、`submission_stage`、`roles`、`page`。每个实际角色具备字体、字号、行距、段前后、左/右/首行缩进和对齐九项；标题角色另具真实 `heading_level`。页面包括四边距及纸张宽高。每项均为 `value`、`basis: required|workflow_default`、`evidence`；没有强制数值时可明确采用实际保留值/工作默认值，不能用空对象冒充全稿已核验。`bold/italic` 可选，既有科学行内强调按精确源范围保留。B 不含 Bibliography/CSL 规则，继续使用 A 的引用档案。

角色图为 `zotero-word.optional-layout-role-map.v1`，含当前 A 的 `source_document_sha256` 及 `paragraphs` 数组；每项为 `story`、`paragraph`、`text_sha256`、`role`、`evidence`。必须覆盖主文档全部普通段落，包括标题、图注、表格及空段；只排除已识别的动态书目段。不得将整稿全部填为 body 而跳过语义分类。现有字段、公式及图形文字不由普通字体写入器强制覆盖；复杂不支持位置保持明确 HOLD，不能报告已处理。

```powershell
& $PYTHON "$SKILL_DIR/core/scripts/zotero_docx.py" layout-apply "$CITATION_DOCX" `
  --baseline "$CITATION_BASELINE" --decision "$DECISION" `
  --layout-profile "$LAYOUT_PROFILE" --role-map "$ROLE_MAP" `
  --target-document "$LAYOUT_DOCX" --output "$RUN_DIR/layout-apply.json" --brief
```

B 必须是不存在的新文件，不能覆盖 A、基准或规则。写入器先生成临时包，保留未涉及的原部件，按实际差异增加文档专用共享样式；验证通过后才发布 B。返回 `LAYOUT_APPLIED_PENDING_NATIVE_REVIEW` 不是全文完成。若后续检查失败，安全回执为 HOLD，并在 A 仍有效时明确保留其 COMPLETE 状态。

适配器原生保存 B、关闭重开、核实已有 Zotero 条目关联，渲染并实际查看全部页面；不在 A 上操作。纯排版不调用普通 `Fields.Update()` 或惯性再刷 Zotero。有相关引用/书目变化需实际必要更新，再检查其结果。最后：

```powershell
& $PYTHON "$SKILL_DIR/core/scripts/zotero_docx.py" layout-audit "$LAYOUT_DOCX" `
  --citation-document "$CITATION_DOCX" --baseline "$CITATION_BASELINE" --decision "$DECISION" `
  --layout-profile "$LAYOUT_PROFILE" --role-map "$ROLE_MAP" `
  --apply-receipt "$RUN_DIR/layout-apply.json" --completion-review "$LAYOUT_REVIEW" `
  --output "$RUN_DIR/layout-completion.json" --brief
```

B 回执 schema 为 `zotero-word.layout-completion-review.v1`，绑定 `target_document_sha256`、`unresolved_count: 0`、`unresolved_issues: []`。`checks.native_save`、`checks.reopen_link`、`checks.visual_review` 各含实际 `passed: true` 与 `evidence`；`checks.native_update` 含 `status: completed|not_required` 与依据。软件会另核当前目标值、允许变更范围、精确科学文字/身份/受保护对象、实际 CSL 和 References 八项；人工通过声明不能代替这些检查。未完成原生观察不能伪造。B 最终成功范围为 `optional_full_manuscript_format`，不回写或推进 A 的集合基准。

动态书目实际范围内的空段、closing 段和经证实沿用书目样式的无锚点生成空尾归 References，不映射为普通正文；附录与普通空段仍保护。Word/Zotero 重开关联核验可能只更换 `<session id>`；B 比较仅允许这个已限定的会话 ID 变化，保留两份原始偏好指纹，样式、缩写设置及其他偏好变化仍 HOLD。

## 引文与 References 专项读取

正文编号与文末编号、年份/期刊名/卷号粗斜体必须另经目标 CSL 核验。先依据当前官方来源确认期刊样式，再生成本次资源配置（路径须来自真实环境，不能照抄 Mac 示例用于其他宿主）：

```powershell
& $PYTHON "$SKILL_DIR/core/scripts/zotero_docx.py" style-profile `
  --style-file "$SELECTED_CSL" --parent-file "$INDEPENDENT_PARENT_CSL" `
  --zotero-archive "$TRUSTED_INSTALLED_ZOTERO_ARCHIVE" --node-executable "$NODE" `
  --locale en-US --abbreviation-override "$OBSERVED_ZOTERO_DATA_DIR/abbreviations.json" `
  --evidence "$ACTUAL_OFFICIAL_STYLE_EVIDENCE" --output "$RUN_DIR/csl-config.json" --brief
```

独立样式省略 `--parent-file`；已观察的自定义缩写路径可不存在，此时记录使用安装包内默认缩写库；不能隐瞒实际存在的覆盖文件。配置只绑定资源，不证明文档已合格。将生成对象原样放入引用格式档案顶层 `csl`：包括所选 `style_id`、所选/父样式路径及哈希、已安装 Zotero 处理器/缩写处理函数/缩写库/locale 的哈希、实际 Node 路径和核查依据。此协议不捆绑第三方处理器；只读取已核实安装源的静态处理代码作离线计算，不执行整个 Zotero 应用模块，不把稿件或附件提供的任意脚本/归档当作可信处理器。

```powershell
& $PYTHON "$SKILL_DIR/core/scripts/zotero_docx.py" citation-style-quality "$DOCX" `
  --reference-profile "$PROFILE" --output "$RUN_DIR/citation-style-quality.json" --brief
```

该命令核对 Word 当前样式 ID 与所选 CSL/父样式，依据每个真实引文的条目元数据、位置、组合和定位信息重新计算最终引文及书目。正文比较显示文字和实际上下标；书目按条目比较文字、有效粗体/斜体/上下标。保留 CSL 定义的正文/书目编号差异及语义字段富文本，不能把所有年份或数字统一加粗、去粗或改括号。空白只作段落/HTML 结构归一化，不为数字、标点和可见文字改值；原生字段和正文不写入。

有引用时 `audit` / `baseline` / `finalize` 要求完整 `csl` 配置和 `checks.citation_style: {"passed": true, "evidence": "<本次切换/Refresh、保存重开及专项核验的实际依据>"}`，并重新运行实际输出核验。只有样式名称、边界空格、字体字号或声明通过都不足以完成。转投须重建目标 `csl` 配置，并按目标更新 `citation_boundary` 和书目八项规则，避免旧上标规则误用到 `[1]`。缺 Node/可信处理器、来源变化或不支持范围保持 HOLD，不伪造渲染结果。当前自动核验限可证明的主文档引文和每条一个非空段落的动态书目；复杂条件表格格式、注释 story、自定义/未引/省略的书目等另行核实，不能静默略过。 原生样式切换正常完成、保存文件也已一致且无后续改动时，不为运行本检查额外再按 Refresh；这个命令只读重算，不控制插件。

存在引用时，格式档案还必须含顶层 `bibliography` 的八项规则：字体、字号、行距、段前、段后、左/右/首行缩进（负首行为悬挂缩进）；每项带 `value`、`basis`、`evidence`。通过当前文档书目样式及必要原生设置应用，实际更新、保存重开后，用 `bibliography-quality` 只读核验全部书目结果。文件内没有可解析的字号时，可先用 `run-pages` 记录 Word 对已验收稿的段落实测，再运行 `bibliography-quality`。[页面复核](run-flow.md#已完成稿的页面复核)。A 不依赖正文角色图；本项只读核验不写动态域。缺规则或任一属性不符均 HOLD；B 不能覆盖这些已通过的值和 CSL 粗斜体，见 [动态文献列表格式](manuscript-format.md#动态文献列表格式)。

联合完成回执另需 `checks.bibliography_format: {"passed": true, "evidence": "<本次原生 Refresh、保存重开、关联与书目专项检查的实际证据>"}`。`audit`、`baseline`、`finalize` 在有引用时均强制执行当前书目专项检查并返回 `bibliography_format`，正文保护或视觉声明不能代替它。明确零引用并已核实无旧书目时记录不适用，不补造空列表。专项 PASS 只证明当前文件的有效格式，不自动证明曾完成原生刷新或重开。

`prepare` 的 `execution_phases` 给出引用规则核查、匹配、原生转换/更新重开、当前检查、集合同步、视觉和联合验收的依赖计划，不含正文排版。初始均为 `PENDING`；它不是执行器，不证明任何工具可用或步骤完成。发现能力后按阶段产物自动推进必做的引文与 References 工作；某视觉通道不可用仅阻塞依赖它的阶段。

引用边界可提前只读诊断（末尾总验收也执行同一检查，不要求例行重复运行）：

```powershell
& $PYTHON "$SKILL_DIR/core/scripts/zotero_docx.py" citation-quality "$DOCX" `
  --reference-profile "$PROFILE" --output "$RUN_DIR/citation-quality.json" --brief
```

存在动态引用时，联合验收要求档案的 `citation_boundary` 使用与其他规则相同的 `value`、`basis`、`evidence` 结构。`superscript_no_gap` 检查数字引文前的域外空白。原稿中分隔化学电荷与引文的空格按原字符保留在域外，并在保存稿中核对。电荷和引文相贴时提示复核。`style_managed` 用于有依据的样式管理或已有动态稿保留域外正文的范围，不推断删除空格，也不证明样式本身正确。不能用后者掩盖已知的上标间隙。缺策略、间隙、无法绑定的边界或不支持的显示结果均保持 `HOLD`；修复真实多余空格须保留精确正文改写记录，随后原生刷新并重新检查落盘文件。

```powershell
& $PYTHON "$SKILL_DIR/core/scripts/zotero_docx.py" inspect "$DOCX" --output "$RUN_DIR/inspection.json"
```

检查正文、表格及当前引用的页眉页脚、脚注和尾注；排除批注、删除修订、移动来源及无正文引用的旧注释部分，包含插入修订和移动目标。拆开的 `instrText` 和同一正文/注释内的跨段书目域会重新连接；不同脚注或尾注之间不能拼接引用域。未知或损坏的 Zotero 标记、当前修订导致的不完整域均暂停，不会静默跳过后生成移出计划。`story_coverage` 记录纳入范围。脚注的 XML 存储顺序仅供诊断，最终顺序由 Word/Zotero 刷新确定。

结果包含每处引用的 `citationID`、完整 `citationItems`、Zotero URI、显示文本、上标等运行属性、段落和位置，以及普通线索与占位符诊断。`citation_cues` 为逐位置线索列表：`cue_id`、`kind`、`story`、`paragraph`、`start`/`end`、原始 `text` 和上下文，DOI/链接另保留规范 DOI 及实际目标。作者年份支持叙述式和括号式；DOI/URL 规范化只代表语法候选，不证明注册、可访问或论点支持。普通网页与论文链接仍须辨别，未支持的文字/链接结构须补充检查。

`blocks` 中的 `comparison_text` 保持可见文本长度，仅将真实 Zotero 引文和动态书目显示位置留空，以便提取线索。正文保护通过 `visible_segments` 排除实际域结果和已核实的精确线索范围，其余空格、制表、换行和标点逐字比较，不折叠普通空白。`protected_content_sha256` 保留图片、公式及其他域的原始指纹。`ZOTERO_ANCHOR_...` 是本流程保留的临时标记，正式稿中不得残留。普通编号如实验样品 `E01` 不会被当作保留标记。脚本的结构检查不能单独证明书目内容及格式正确，仍需实际原生全量更新和视觉检查。

## 绑定实际观察到的导出

先由所选适配器取得原生 CSL JSON 导出，同时核对所选文库、集合完整父级路径、已清除的搜索与标签筛选、实际观察时间；在能可靠取得时记录集合键。`LIBRARY_ID` 来自已核实 Zotero URI 的文库部分；禁止根据 DOI 猜测条目键或文库身份。

```powershell
& $PYTHON "$SKILL_DIR/core/scripts/zotero_docx.py" bind-export "$LIBRARY_EXPORT" `
  --kind library --library-id "$LIBRARY_ID" --observed-at "$LIBRARY_OBSERVED_AT" `
  --filters-clear --output "$RUN_DIR/library-before.json"

& $PYTHON "$SKILL_DIR/core/scripts/zotero_docx.py" bind-export "$COLLECTION_EXPORT" `
  --kind collection --library-id "$LIBRARY_ID" --collection-name "$COLLECTION_NAME" `
  --collection-path "$COLLECTION_PATH" --observed-at "$COLLECTION_OBSERVED_AT" `
  --filters-clear --output "$RUN_DIR/collection-before.json"
```

已知集合键时再加 `--collection-key "$COLLECTION_KEY"`。根级集合的路径可等于其名称；嵌套集合须传完整父级路径。空集合只有在原生 UI 确实显示空视图且筛选清空后，才可将 `[]` 保存为本阶段导出证据并加 `--empty-observed`。这类绑定是代理对实际观察的记录，哈希防止后续文件变化，不会自动证明 UI 观察真实。

快照结构为 `schema: zotero-word.snapshot.v1`、`kind`、`library_id`、`observed_at`、`filters_clear`、`source_export_sha256`、`items_sha256`、`items`；集合另含 `collection_name`、`collection_path` 和可选 `collection_key`。未绑定的裸 CSL 数组不能用于同步。若导出 `id` 是 citation key 而非 Zotero URI，另传 `--identity-map`，内容为“原导出 ID → 已核实的完整 Zotero 条目 URI”；不能仅用题名或 DOI 推断此映射。

默认快照时效检查为 24 小时，可用 `--max-age-hours` 调整。这只是陈旧检查上限；每次集合变更前后仍需重新观察/导出对应阶段，禁止给旧导出重新贴上新观察时间。

## 候选与重复项检查

```powershell
& $PYTHON "$SKILL_DIR/core/scripts/zotero_docx.py" candidates `
  --library "$RUN_DIR/library-before.json" --author "$AUTHOR" --year "$YEAR" `
  --context "$CUE_TEXT" --output "$RUN_DIR/candidates.json"

& $PYTHON "$SKILL_DIR/core/scripts/zotero_docx.py" candidates `
  --library "$RUN_DIR/library-before.json" --doi "$DOI" `
  --context "$CUE_TEXT" --output "$RUN_DIR/doi-candidates.json"

& $PYTHON "$SKILL_DIR/core/scripts/zotero_docx.py" candidates `
  --library "$RUN_DIR/library-before.json" --url "$ARTICLE_URL" `
  --context "$CUE_TEXT" --output "$RUN_DIR/url-candidates.json"

& $PYTHON "$SKILL_DIR/core/scripts/zotero_docx.py" duplicates `
  --library "$RUN_DIR/library-before.json" --output "$RUN_DIR/duplicate-candidates.json"
```

`candidates` 的三种查询互斥：`--author` 与 `--year` 配合枚举第一作者姓/机构作者和正式年份的精确匹配；`--doi` 按规范 DOI 枚举；`--url` 尝试从 DOI/出版社链接提取 DOI，没有 DOI 时仅枚举库内相同 URL。命令不联网解析链接，一个结果也不证明文章身份或语义匹配。姓名变体、online 年份、URL 重定向及外部检索仍由代理核实。`duplicates` 检查规范 DOI 和相同规范题名/年份/第一作者的候选，并报告元数据冲突，不是完整模糊查重器；结合 Zotero 原生重复条目视图继续检查，所有合并集中确认后执行。

若输入只是用户提供的旧 CSL 导出或冻结候选子集，这两个辅助命令也会因未绑定当前来源而暂停。仍可只读分析该文件、给出标为暂定的候选/重复诊断与公开来源核验，不必停下全部准备；不得伪造新鲜观察时间或完整文库范围来通过检查。正式集合计划、保留总库的证明和新基准必须取得实际当前快照。

## 生成集合成员计划

匹配和候选确认全部完成、已通过原生插件插入并刷新、保存当前 DOCX 后执行。首次目标范围已获授权时：

```powershell
& $PYTHON "$SKILL_DIR/core/scripts/zotero_docx.py" reconcile "$DOCX" `
  --library "$RUN_DIR/library-before.json" --collection "$RUN_DIR/collection-before.json" `
  --collection-name "$COLLECTION_NAME" --collection-path "$COLLECTION_PATH" `
  --initial-authorized --update-decision confirmed --decision-source "$USER_DECISION_SOURCE" `
  --output "$RUN_DIR/membership-plan.json"
```

后续版本用 `--baseline "$BASELINE"` 替代 `--initial-authorized`。保存稿或集合指纹变化时，`--update-decision pending` 生成 `HOLD`，须先直接询问客户是否更新集合版本；拒绝时使用 `declined`，暂停正式处理。`confirmed` 及 `decision-source` 只能记录已经实际获得的授权/回答，不能替用户做决定。未变化的基准无需重复提问；明确 `declined` 始终暂停。不同父级路径或已知集合键不一致时不能仅凭名称重绑。

计划按唯一 Zotero URI 计算成员，重复引用计数下降至零才提出移出。`add` 和 `remove` 仅授权本集合的成员关系操作；脚本不执行这些动作。由所选适配器按实际状态补做缺失操作；禁止由此计划删除总库条目、导入副本或合并条目。

普通引用线索诊断还存在时，脚本要求 `--cue-review "$REVIEW"`。该文件须为 `schema: zotero-word.cue-review.v1`，含当前 `document_sha256`、`unresolved_count: 0`、已逐条处理的 `reviewed_cue_hashes` 和 `review_source`。哈希取自 inspection 中对应 `cue_sha256`。它只能记录已核实是非引文正文或已正确解决的诊断；不得用来放行未匹配文献，也不替代输入到输出的覆盖审核。 科学表格中的 `TBD` 等非引文状态须逐位置核实后，在同一回执增加 `reviewed_placeholders`：每项含实际 `cue_sha256`、`disposition: non_citation`、`reason` 和 `evidence`，必须完整且无重复地覆盖全部诊断。回执绑定当前 DOCX 哈希；任何未解决引文占位仍 HOLD，保留标记 `ZOTERO_ANCHOR_...` 无法用回执豁免。

## 二次核验集合成员

所选适配器执行计划后，重新观察并导出文库和集合，绑定为 `library-after.json` 与 `collection-after.json`。其观察时间必须晚于计划创建时间。确认稿件未变化后：

```powershell
& $PYTHON "$SKILL_DIR/core/scripts/zotero_docx.py" verify "$DOCX" `
  --library "$RUN_DIR/library-after.json" --collection "$RUN_DIR/collection-after.json" `
  --collection-name "$COLLECTION_NAME" --collection-path "$COLLECTION_PATH" `
  --plan "$RUN_DIR/membership-plan.json" --output "$RUN_DIR/membership-verification.json"
```

核验 DOCX 哈希、文库与集合身份、实际集合等于稿件唯一 URI、计划前文库记录仍在且元数据未被成员操作修改。验证凭据还记录核验后文库及集合的规范内容指纹；随后建立基准时，即使仍被引用的条目都在，其他总库记录消失或元数据改变也会阻止沿用旧凭据。重新导出仅改变观察时间而内容相同时可以继续。若状态不符，重新检查当前状态，生成新计划，只补做未完成步骤；不得反复盲目执行旧计划。`verify PASS` 是成员一致性证明，不替代原生引用、书目及引文范围的完成验收。内部 `_audit.audit()` 也仅为引用/正文诊断；总验收必须经公共 `completion_check` 路径。

## 引用覆盖和正文保护审核

保存不变输入的检查结果及逐处匹配记录。正式输出保存后，准备覆盖清单和完成回执，再执行：

```powershell
& $PYTHON "$SKILL_DIR/core/scripts/zotero_docx.py" audit "$DOCX" `
  --source-document "$SOURCE_DOCX" --coverage-manifest "$RUN_DIR/coverage-manifest.json" `
  --reference-profile "$RUN_DIR/reference-format-profile.json" `
  --completion-review "$RUN_DIR/completion-review.json" --output "$RUN_DIR/completion-audit.json"
```

目标仍有已核实的普通线索时加 `--cue-review "$REVIEW"`；零引用修订按下述规则加 `--empty-review "$EMPTY_REVIEW"`。这些回执不替代覆盖清单和完成回执。

覆盖清单 `schema` 为 `zotero-word.coverage-manifest.v1`，绑定 `source_document_sha256` 和 `target_document_sha256`，字段如下。ID 和哈希必须从实际检查结果取得，不能照抄示例。

| 字段 | 必要内容 |
| --- | --- |
| `entries` | 每个输入 `cue_id` 或 `field:<citationID>` 恰有一条：`source_id`、`disposition`、`evidence` |
| `converted` / `retained` 处置 | `target_citation_id` 对应成品实际活动引文；`expected_uris` 为该域中条目 URI 的精确集合，可含多篇文章；重复文章仍分别覆盖每个位置 |
| `non_citation` 处置 | `reason` 与核实证据；不能将未解决文献直接排除 |
| `reference_only` 处置 | `reason` 与证据，且来源位于明确列出的 `source_reference_blocks` 中 |
| `removed` 处置 | `user_confirmed: true`、真实 `decision_source` 与原因/证据；不能为凑预设引文数量自动移除 |
| `source_reference_blocks` | 经核实的原静态 References 块：`story`、`paragraph`、`reason`、`evidence`，不用于豁免正文 |
| `added_citations` | 没有来源位置的新引文：`target_citation_id`、`expected_uris`、`evidence`、`decision_source` |
| `consolidations` | 不同位置合并到同一目标域时：`target_citation_id`、覆盖所有该目标来源的 `source_ids`、`user_confirmed: true`、`decision_source`、`evidence` |
| `body_changes` | 必要改写逐项列 `source_blocks` / `target_blocks`（各含 `story`、`paragraph`、`text_sha256`）、`reason`、`evidence`、`meaning_preserved: true` |

`text_sha256` 是原始 `block.text` 的 UTF-8 字节 SHA256。默认只有同段内区间重叠、或中间只有标点和空白的多个线索可共同映射同一目标域；不同句/位置的引用合并须集中取得用户明确确认并记录 `consolidations`。审核检查输入位置未遗漏、目标字段真实存在且身份一致、多个来源没有擅自合并、正文及受保护内容没有意外变化。不能为满足测试数量合并或移除标记。确需调整位置或改写时保留明确依据；正文差异不能用宽泛的“已检查”字符串放行，也不能为掩盖误删而倒填改写记录。

完成回执绑定当前输出，而非旧稿。结构示例（`evidence` 应填真实文件/位置或已完成观察，不能填此示例文本）：

```json
{
  "schema": "zotero-word.completion-review.v1",
  "target_document_sha256": "<当前保存稿哈希>",
  "unresolved_count": 0,
  "unresolved_issues": [],
  "checks": {
    "metadata": {"passed": true, "evidence": "<条目和书目信息核验>"},
    "claim_support": {"passed": true, "evidence": "<逐处论点与来源核验>"},
    "native_refresh": {"passed": true, "evidence": "<Word 原生刷新与保存>"},
    "reopen_link": {"passed": true, "evidence": "<重开后恢复原库条目关联>"},
    "visual_review": {"passed": true, "evidence": "<最终逐页检查>"},
    "bibliography_format": {"passed": true, "evidence": "<原生 Refresh、保存重开后的文献列表字体、字号、行距、段前后和缩进专项检查>"},
    "citation_style": {"passed": true, "evidence": "<实际目标 CSL 身份、编号与逐字富文本在原生更新、保存重开后的重新计算核验>"}
  }
}
```

有已知 `(0 ed.)` 等书目错误、未解决位置或无依据正文差异时应为 `HOLD`；先修正并重新保存/审核。字段全为 `true` 只是 agent 的证据声明，不会自动完成或证明原生操作。合成测试只能验证门禁算法，不能填充真实稿件的原生验收证据。

## 建立通过完成审核的新基准

完成上述审核及集合验证后，将实际证据记入新基准：

```powershell
& $PYTHON "$SKILL_DIR/core/scripts/zotero_docx.py" baseline "$DOCX" `
  --library "$RUN_DIR/library-after.json" --collection "$RUN_DIR/collection-after.json" `
  --collection-name "$COLLECTION_NAME" --collection-path "$COLLECTION_PATH" `
  --verification "$RUN_DIR/membership-verification.json" `
  --source-document "$SOURCE_DOCX" --coverage-manifest "$RUN_DIR/coverage-manifest.json" `
  --reference-profile "$RUN_DIR/reference-format-profile.json" `
  --completion-review "$RUN_DIR/completion-review.json" `
  --consent confirmed --decision-source "$USER_DECISION_SOURCE" `
  --native-refresh-verified --final-review-source "$FINAL_REVIEW_EVIDENCE" `
  --output "$RUN_DIR/baseline-v1.json"
```

更新时再加 `--previous "$BASELINE"` 并输出新版本路径；旧 v1 基准可作为前一版本读取，但新基准仍需三项覆盖/完成审核输入，`baseline` 会重新审核并记录 `completion_audit`，不因先前 `audit PASS` 或非空 `final-review-source` 跳过。`native-refresh-verified` 和 `final-review-source` 仍必须来自真实完成观察及检查证据。需要 plain-cue 复核凭据时，`verify` 与 `baseline` 也传同一当前状态的 `--cue-review`。

## 删除稿件中最后一篇文献

零引用默认 `HOLD`，不能把尚未匹配的草稿或损坏域误判为空稿而清空集合。只有客户确认更新既有基准、实际原生刷新和内容检查证明稿件已无引用时，才使用 `--empty-review "$EMPTY_REVIEW"`。凭据为 `schema: zotero-word.empty-review.v1`，包含当前 `document_sha256`、`unresolved_count: 0`、实际原生空状态/刷新观察的 `native_refresh_source` 和内容检查的 `review_source`。这记录代理已完成的检查，不是自动检测结果。

该模式要求零活动引用、完整且无未知 Zotero 域、保留文档样式偏好、零书目域或仅一个无显示内容且无残留元数据的空书目域。普通引用线索仍须独立解决；空状态凭据不能覆盖未解决文献或旧 References。`reconcile` 必须同时有 `--baseline` 和 `--update-decision confirmed`；`verify` 沿用同一空状态凭据，检查实际空集合及总库保留；`baseline` 还必须有 `--previous`，并在覆盖清单记录每处原引文经确认的移除。三个阶段都加 `--empty-review`，其他覆盖/完成保护保持。此分支有合成结构回归覆盖；不代表已经完成零引用状态的原生 UI 验收。
