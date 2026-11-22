# 引文与 References 格式

本页用于核对引文边界与动态文献列表的专项格式。常规引文任务完成后交付已验稿件。用户另行委托全文排版时，再按[可选全文格式](optional-layout.md)从已验稿创建副本。

CSL负责正文引文和文献著录顺序、标点、编号及行内富文本；Word的Bibliography样式和实际属性影响列表字体、字号、间距及缩进。两者分别核验。只切CSL不能代替书目八项外观检查；已合格值保持，不为产生改动而重设。

## 引用格式档案

使用小型档案，schema为 zotero-word.reference-format-profile.v1。记录期刊、适用文章/投稿阶段、bibliography 八项规则、citation_boundary 和真实 csl 配置；不含 roles、page、源/最终全稿角色图。只读实现为 core/scripts/reference_format.py，公共命令使用 --reference-profile；旧 --format-profile 参数名仅作 A 的兼容别名，不能替代 B 独立的排版档案或同意门禁。具体结构见 [命令说明](script-commands.md)。

规则逐项包含 value、basis（required或workflow_default）和evidence。先查当前官方引用指南及所选CSL；指南没有固定字体/字号/缩进时保留清晰现值或明确选工作默认值，不能称为期刊强制数值，也不能因缺固定值停在“待查具体数值”。记录核查日期、官方地址、实际CSL及父样式哈希和阶段。发表PDF外观不能充当投稿稿件规范；样式的一种正确输出不等于自由格式投稿的唯一选择。

通过 style-profile 重新绑定实际来源，生成的是csl配置，不是引用格式通过回执。不能复制他稿配置或旧完成结论来替代当前实际来源和结果。

## 动态文献列表格式

有有效引用时，bibliography完整声明以下八项：

- font_family、font_size_pt；
- line_spacing；
- space_before_pt、space_after_pt；
- left_indent_pt、right_indent_pt、first_line_indent_pt。

长度用磅；负 first_line_indent_pt 表示悬挂。line_spacing 的value为 mode与value对象，multiple用倍数，exact/at_least用磅值。条目间距同时核对段前/后，不能只写“同正文”。有据核实零引用且无旧书目的特殊状态不补造空列表规则。

通过当前可用且已验证的Word通道，修改**当前稿件**Bibliography样式及有依据的书目属性，不改Normal模板或“基于此模板的新文档”。只改确有差异的属性，保留CSL的年份粗体、期刊/卷斜体、上下标、条目文本及原生域身份。禁止清除全部格式、解除域、手打编号或覆写成静态列表。A 中已有 References 标题不重设字体或大纲；新生成书目需要新增标题时，只新增辨认该列表所需的文字并记录。

本批引用/元数据/样式/书目变更后完成一次必要原生更新；样式转换已真实更新全稿时不机械叠加Refresh。原生保存、关闭重开并核对关联后，检查同一落盘状态：

    "$PYTHON" "$SKILL_DIR/core/scripts/zotero_docx.py" bibliography-quality "$FINAL_DOCX" \
      --reference-profile "$RUN_DIR/reference-format-profile.json" --output "$RUN_DIR/bibliography-quality.json"

只读检查解析实际有效字体、字号、行距、段前后和三项缩进。对完成回执绑定的交付文件，如果仅剩普通文字的默认字号无法解析，检查会向 Mac Word 读取整段书目域的实际字号，并核对文件指纹和书目全文。Word 窗口含有未保存的编辑时，只读取与已验收磁盘文件逐字节相同的临时只读副本。混合字号、原生读数缺失或与规则不符时保留待核对状态。公共完成核验须有当前书目外观的真实证据及机械检查通过。若最后检查不符，不能只改档案目标或直接强刷外观后省略必要原生更新。

刷新重设格式时，诊断Word样式、直接属性与CSL的实际所有者，再做一次有依据恢复，重复“必要更新→保存重开→当前检查”；仍失败则HOLD。中西文字体分别解析；字符单位缩进、无法确定的继承/主题字体或冲突缩进不猜。正确的CSL字重/斜体由 [CSL输出专项](journal-styles.md#csl-输出专项核验) 逐字比较，八项外观不声明“全部普通体”。

## 原生引文边界

citation_boundary同样含value、basis、evidence，按实际请求和核实规则选择：

- style_managed：已有动态引文的纯转投默认保留域外普通正文/空格；字段内部文本与上标由真实CSL另行强制核验。明确记录目标样式及保持正文的范围依据，不能把此值当作免做CSL检查。
- superscript_no_gap：用户明确要求清理数字上标前间距，或当前期刊有明确要求时使用。只检查完整数字或数字组与能明确定位的域外空白。化学电荷与编号之间原有的普通、窄或不换行空格应保留在域外。电荷与编号相贴时，对照原文复核。制表符、换行和字段内的空白须单独判断。

不把某次ES&T空格问题推广为全部稿件强制去空格。引用直接相关、依据明确的间距/标点/叙述调整仍属于本skill；逐处最小编辑并记录精确 source/target 文本哈希、位置、原因和保留原意的证据，交由body_changes及正文保护核验。A 中其余正文不润色、不改字体或版式；B 只在其独立同意与允许变化计划下处理格式，不改科学文字。可用只读诊断：

    "$PYTHON" "$SKILL_DIR/core/scripts/zotero_docx.py" citation-quality "$FINAL_DOCX" \
      --reference-profile "$RUN_DIR/reference-format-profile.json" --output "$RUN_DIR/citation-quality.json"

边界通过不证明上标、括号和组合编号正确；这些须由真实CSL和实际保存结果核验。相关文字/字段发生变更后仍须最终原生更新和保存重开；只有只读检查没有变化时不追加原生更新。

## 视觉与正文保护

在真实图片/页面中查看全部实际引文及完整References，检查编号、换行、字重、缩进、条目间距和与邻近图文的遮挡。A 需要时查看全文页图来覆盖所有引用，但不借此提前执行 B。保留标题页空白、图表大小及已有分页；只记录引用可读性与交付相关问题，不把未处理的正文美化视为引用任务未完成。

普通正文、图片、数学公式、其他域和表格结构仍须保护。引用线索转换和有依据的引文边界修改精确列入变化清单；不得用整段“已审核”豁免掩盖无关改写。Word重存规范化也须按当前工具支持和实际证据处理，不能删原始证据或伪造通过。

B 不能用全选正文/样式重设覆盖已正确的 References。B 最终保存重开后继续核对本页八项实际值及 CSL 富文本；只在确实影响引用/书目时进行必要原生更新，不无条件重复 Refresh。A 原始完成回执保留，B 另存结果，不能将有意的 B 格式变化送回 A 的“版式不变”比较来误报或借此关闭保护。
