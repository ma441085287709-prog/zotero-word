# 目标期刊格式与转投

本页属于通用核心；需要操作 Word、Zotero 或外部资源的步骤，由所选平台适配器使用其真实可用工具执行。

用户给定期刊后，先确定完整刊名和稿件阶段，再查其当前官方作者指南、官方/维护方 CSL 及本机实际安装样式。不要仅凭 CEJ、JHM、EST 缩写、旧教程或样例文章推定规则。记录检索日期、页面地址、适用阶段、CSL `id`、更新时间及本地文件哈希；依赖样式还要记录其独立父样式。

CSL 控制引文、文献著录结构及其支持的书目布局；动态文献列表的字体、字号等还须通过本稿 Word 的 Bibliography 样式与实际属性处理。阶段 A 只整理这些位置中的引用和动态文献列表，保持其余正文及版式；A 完成后，用户明确选择 B 才核查和应用全文期刊格式。文献列表外观见 [引文与 References 格式](manuscript-format.md)。A 按引用规范适用的投稿阶段核查，不提前制作全文规则/角色图。B 可复用同一官方指南中已核实的来源，但只在明确同意后展开全文格式准备；不能照发表 PDF 推定投稿稿件排版；未规定的书目字体等可保留清晰现值或标为工作默认值，不称为强制规范。

## 选择与差异处理

1. 从出版商官方期刊页进入作者指南，确认正文标记、排序、作者列表、期刊缩写、文章号/页码、DOI 和初投例外。指南无法访问时明确缺口，继续核对可用官方 CSL，不把“样式输出正确”写成“全部投稿细则核验完成”。
2. 在 Zotero 的样式管理/“获取更多样式”中查找完整刊名，核对其 `id` 与指南对应关系。安装已有官方样式属于被请求格式处理的正常步骤；遇到新增权限或来源异常则按实际授权边界处理。
3. 在 Word 的 **Document Preferences** 选择确切期刊；同一出版社或父样式的通用名称只能在确认等价时使用。不要因本机已有某通用样式就跳过期刊核对。
4. 官方指南与 CSL 有明确、可唯一解释的差异时，自动制作本稿专用 CSL 副本修正；不修改共享原样式，不修改论文元数据来迁就格式。规则相互冲突、阶段不明或有多种合理解释时，集中列出差异、证据和选项供用户确认，然后继续。

## CSL 输出专项核验

正文引文的方括号、圆括号、上标和组合/范围，与 References 的编号前后缀及富文本分别由真实 CSL 计算；不得从正文 `[1]` 推定列表也应为 `[1]`，或把所有数字样式统一处理。选定期刊后，用 `style-profile` 在验证来源后生成引用格式档案的顶层 `csl` 配置；它只生成配置，不证明稿件已改好。具体参数见 [命令说明](script-commands.md)。记录选中样式 ID、样式文件及独立父样式（当前支持单一独立父样式，如有）、实际 Zotero 安装包及其中处理器/缩写处理器、当前缩写数据库与语言包的来源及哈希、实际可用 Node 路径、语言和依据；显式缩写覆盖也须记录文件与哈希。不要凭模型记忆手填一套期刊预期文本。

只读 `citation-style-quality` 使用已安装 Zotero 的真实静态处理器和缩写来源进行本地计算，从当前保存稿的真实引用元数据生成预期输出，再与落盘结果比较：文档偏好中的选中 CSL ID 须精确匹配配置；依赖父样式、语言及自动缩写状态须一致；正文逐处比较显示文本与实际上标/基线；书目比较各条编号标签、完整文本和逐字有效加粗、斜体及垂直位置。组合引用和范围由处理器计算，不自行拼接编号。样式继承和直接属性均须解析，不能只看运行上是否写了 `b` / `i`。

这项计算不控制 Zotero/Word、不制造或改写引用域，也不把处理器代码复制进 skill。先发现本机实际 Node 与来源文件；缺能力或无法验证来源时，该检查保持 HOLD，继续其他独立工作，不能泛称整个流程均不可用。当前核验限于可明确绑定的主文档活动引用及完整原生书目；其他 story、不支持的文献类型、自定义/增补/省略的书目等须报告具体不支持项，不能当成已通过或以样例输出替代。仅经稿件状态流程确认的当前零引用且无旧书目状态可不适用。

必要的原生全量更新（包括已完成的样式转换）、保存、关闭重开后对同一最终状态执行：

```powershell
& $PYTHON "$SKILL_DIR/core/scripts/zotero_docx.py" citation-style-quality "$FINAL_DOCX" `
  --reference-profile "$RUN_DIR/reference-format-profile.json" --output "$RUN_DIR/citation-style-quality.json"
```

公共联合验收须同时有 `checks.citation_style` 的 `passed: true` 与实际证据，并由当前输出的机械比较通过；只声明“已选择期刊”、边界无空格或八项书目外观合格均不足。此核验不替代当前引用规范、文献身份与必要论点检查、原生更新或引用相关视觉检查。

## 年份加粗与转投残留

对使用本次已核实 ACS 父样式的期刊论文，年份加粗、期刊及卷号斜体属于样式输出，不能一律去除。依据应是处理器产生的字段与逐字格式；DOI 中出现的 `2011` 等数字不是发表年份，禁止以所有四位数字为年份进行替换。不同文献类型的规则可不同。

发现多余加粗/斜体、缺失预期格式或转投后混合编号时，先判断目标样式、CSL 行内输出与 Word 基础/字符/直接属性的实际来源。通过真实Word的Document Preferences选择正确样式并完成必要的原生全量更新；只有证实本稿 Bibliography 基础加粗/斜体是额外残留时，才恢复该基础属性，保留 CSL 规定的行内富文本。不得对整份书目统一取消加粗/斜体、文本替换括号或手打编号。实际修复并入本批必要原生更新，再保存、关闭重开并复验CSL输出与八项书目外观；失败按有界恢复规则保持 HOLD。

## 本稿专用 CSL

只在需要时创建。若期刊 CSL 是依赖样式，先读取它声明的 `independent-parent`，复制独立父样式形成可编辑的本稿样式；不能只给依赖样式增加 `citation`/`bibliography` 段落。保留源文件与许可证/来源信息，设置独特的 `title`、`id` 和 `rel="self"`（例如带本稿稳定标识的名称），避免覆盖原期刊样式。只修改有官方证据支持的规则，记录原规则、新规则、来源和前后哈希。

通过 Zotero 原生样式管理安装本稿副本，并在本稿的 Document Preferences 选择它。先用包含多作者、DOI、页码或文章号、组合引用、重复引用的独立样例检查预览，再对稿件完成一次必要的原生更新。核验真实编号、上标和书目内容；XML 合法或安装成功本身不构成格式通过。若效果异常且已有恢复依据，保留诊断副本后切回已记录样式，按实际需要完成原生更新。输出时交付适用的本稿 CSL，后续仍可重现；不能承诺独有样式在未安装的其他电脑上自动可用。

## 转投的执行与验收

转投以同一份活动引用文档为基础，通过 Document Preferences 全局切换目标期刊并观察全量更新完成，再原生保存、读取输出核验。成功的样式转换已实际更新全稿时不叠加Refresh；仅在仍有延后/未更新结果或之后又有相关修改时补一次必要更新。不要逐条重写引文数字、解除域或重建静态 References。仅样式切换不应增删集合成员。

转投同时更新 [动态文献列表八项规则](manuscript-format.md#动态文献列表格式)，在当前文档应用相应 Bibliography 样式属性，并核对 CSL 对行距、条目间距和缩进的影响。保留期刊所需斜体、加粗和上下标；必要原生更新、保存重开后检查实际有效格式。仅 CSL 文本输出正确或视觉回执通过，不能替代书目专项核验；刷新重设时按核心恢复规则诊断，不以末尾直接格式覆盖绕过复验。

更新本稿 `csl` 来源配置并运行 [CSL 输出专项核验](#csl-输出专项核验)，分别检查正文与列表编号、上标和各条富文本，发现旧样式残留。ES&T→其他期刊→ES&T 的往返验收须保留每一轮真实来源和保存结果，不能仅比对一个数字或把计划中的原生测试写成已验证。

保存切换前后引用的 Zotero URI、条目元数据、出现次数和引用域身份；没有引用内容编辑时，这些身份应保持。正文普通语句、表格、图片、公式和独立图表编号也应保持。对比实际书目文本及富文本：两个数字期刊的单个 `[1]` 可能外观相同，因此还需检查作者顺序、年份位置、期刊缩写和 DOI。上标须验证真实格式属性（Word/OOXML 的 `vertAlign=superscript`），不是把字符换成上标数字字形。

保存、关闭、重开后验证现有条目关联和最终结果。只读查看/取消未改的对话框不触发额外全稿Refresh。若用户要求切回原刊，核对恢复后的格式，不要求整份 DOCX 二进制相同，因为 Word 可能规范化字体、语言或内部样式 ID。

## 核查时区分投稿阶段

ES&T 官方指南在 2026-11-10 核查时更新于 2026-10-05，并允许初投用任意完整的引用格式、包括文章标题。选择期刊 CSL 的一种输出不能泛化为初投唯一可接受格式；新任务仍应读取当前要求。[ES&T 作者指南](https://researcher-resources.acs.org/publish/author_guidelines?coden=esthag)

官方完整指南返回 403 等访问限制时，不要据此推定旧教程正确，或声称全部投稿要求已查清。只陈述本次真正核对的规范与样式。各适配器独立保存自身的历史版本和验收记录，不能把这些记录当作新任务的当前期刊要求；自定义 CSL 的首次使用须单独验证。

## 官方入口与维护方源码

核查日期：2026-11-10。页面与样式会变化，以下是检索入口，不是冻结的期刊规则副本。

- [Zotero 样式库](https://www.zotero.org/styles)：原生获取或安装期刊样式。
- [Zotero Word 插件](https://www.zotero.org/support/word_processor_plugin_usage)：Document Preferences 和 Refresh 的作用。
- [CEJ 官方作者指南](https://www.sciencedirect.com/journal/chemical-engineering-journal/publish/guide-for-authors)、[CEJ 样式入口](https://www.zotero.org/styles/chemical-engineering-journal)、[CEJ 依赖定义](https://raw.githubusercontent.com/citation-style-language/styles/master/dependent/chemical-engineering-journal.csl)。
- [JHM 官方作者指南](https://www.sciencedirect.com/journal/journal-of-hazardous-materials/publish/guide-for-authors)、[JHM 样式入口](https://www.zotero.org/styles/journal-of-hazardous-materials)、[JHM 维护方源码](https://raw.githubusercontent.com/citation-style-language/styles/master/journal-of-hazardous-materials.csl)。
- [ES&T 样式入口](https://www.zotero.org/styles/environmental-science-and-technology)、[ES&T 依赖定义](https://raw.githubusercontent.com/citation-style-language/styles/master/dependent/environmental-science-and-technology.csl)、[ACS 父样式源码](https://raw.githubusercontent.com/citation-style-language/styles/master/american-chemical-society.csl)。

`zotero.org/styles/<id>` 可返回 CSL XML 而非网页；浏览工具显示不支持内容类型时，通过 Zotero 原生安装/导出读取，或读取上列维护方原始文件。不要把渲染器拒绝 XML 当作样式不存在。

## 可选全文期刊格式

仅当用户另外委托全文匹配时，以已验引文稿启动[可选全文格式](optional-layout.md)。分别核查标题、摘要、各级标题、正文、表题/表注/表内文字、图注及页面等适用规则，按稿件阶段记录官方来源、日期与实际要求。期刊未规定具体数值时保留其允许值或明确工作默认值，不能以出版版式充当强制规范，也不为缺固定值把已授权 B 停留在建议。

B 从 A 新建排版副本，允许计划内格式变化并保护内容和真实字段；A 的引用完成状态保留。纯排版不重新切 CSL、同步集合或无故刷新。B 保存重开后仍需复核本页 CSL 及 References 八项、恢复已有条目关联并查看全文实际布局；只有这些和 B 的允许格式检查通过，才报告全文格式匹配完成。
