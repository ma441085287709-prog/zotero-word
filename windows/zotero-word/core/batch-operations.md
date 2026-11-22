# 批量检索、导入与集合操作

批量阶段减少的是库检索、元数据获取、导入和集合管理的工具往返；不替换 Word 原生 Zotero 插件，不保证 Zotero 全量刷新更快。先发现宿主真实 MCP、CLI、本地 API 或原生工具能力。可用批量接口则优先使用，遵守服务端实际批次上限；不可用则继续现有原生适配，不安装未知工具或扩大权限。

先用 `library-health --config …` 检查已配置的实际接口、版本、访问范围和库身份；配置文件存在或 connector ping 成功不代表库可访问，更不代表可以写。`batch-execute` 才是受限执行入口，具体配置、授权、日志与下一动作见 [库操作执行](library-execution.md)。本轮实现的 HTTP 通道使用现有 Zotero Local API；主机工具请求通道只输出待派发白名单请求，宿主实际调用后还须重新观察。没有真实可用能力时明确返回 `CAPABILITY_MISSING` 并转现有原生 UI，不静默返回空库或宣称批次已完成。

设计参考 [54yyyu 的 CLI](https://github.com/54yyyu/zotero-mcp/blob/main/docs/cli.md) 和 [Xevos117 的批量工具](https://github.com/Xevos117/mcp-zotero/blob/main/README.md) 的结构化调用方式；本包代码为独立实现，不捆绑、自动安装或运行这些项目。它们的字段注入器不进入当前 Word 动态引用通道。

## 检索与导入

1. 一次 `prepare --brief` 得到逐位置线索、去重检索组及已有活动域。重复 DOI/URL/作者年份只复用检索结果，不合并引用位置或跳过各处语义核验。已有活动域继续保留复核。
2. `batch-plan` 把尚需公开检索的组分批列出；已在完整库快照中枚举到候选的组进入语义复核清单，不重复发送库搜索。输出是工具任务，不表示已检索、已选中或已导入。
3. 身份及上下文核实后，整理 `--import-requests`。每项含实际 `source_ids`、DOI 或 URL、`identity_verified: true` 和具体 `evidence`；可携带已核实 `metadata`。确定 DOI 必须与该位置的原始检索组一致，冲突保持 HOLD，不能只凭 `identity_verified` 替换为其他论文。普通出版社 URL 经公开元数据核验后可按新 DOI 导入，但须在顶层 `URL`、`metadata.URL` 或明确的 `resolved_from_URL` 保留与原始 URL 的规范化绑定；缺该绑定仍 HOLD，计划将真实来源保存在 `resolved_from_URLs`。作者年份线索可经语义核实后映射新 DOI。只导入确实缺少的条目。相同标识的请求合并来源位置，但元数据或选择冲突必须先解决。
4. 导入前再次检查目标库：单一且已核实的现有身份复用；多个现有身份集中确认，不自动选最高分、不自动 Merge。明确选项可记录 `selected_uri`、`selection_confirmed: true` 和实际 `decision_source`。无 DOI/文章 URL 的记录沿用已核实的原生导入路径，不用标题猜身份。
5. 批量读取导入后的真实条目，保存实际 URI、返回版本及成功/失败/未知结果。未知写入先重查；报错或超时不等于没有写入。缺 PDF 后置处理，不阻断已解决引用。

`core/scripts/library_batch.py` 仍只输出 JSON 计划和重查结果；`core/scripts/library_executor.py` 执行其实际支持且已具备访问的动作。`idempotency_key` 是本地稳定的工作标识，不证明外部服务具有幂等导入能力。适配器须实际发现该能力；没有服务端保证时，使用导入前检索与导入后观察保护，不能盲目重放未知写入。

## 集合同步

先由现有 `reconcile` 生成用户已确认的 `READY` 差集。批次仅包含 `add_membership` 和 `remove_membership`，不创建/重建集合，不删总库条目、不移入废纸篓、不合并记录。删最后一次引用后只移出当前稿件集合，其他集合保持。

导入准备通常绑定原稿；正式集合差集绑定已转换保存稿，二者哈希可能不同。转换后创建新的 membership-only 批次即可，不把不同版本强绑在同一计划。集合标识沿用实际 library ID、key、名称和完整路径；API 需要 key 而当前证据没有 key 时，先核实真实 key，不能凭同名重建集合。

写入前核对条目与集合状态，工具提供版本/条件更新时使用真实版本；缺此能力不伪造版本。批次结束或中断后，重新导出/观察库及集合，用 `batch-status` 重查已完成部分。只有计划内成员变更已在更新后的集合中实现、普通元数据不变、版本正常推进且已暴露的其他集合保持时，才接受相应版本变化。既有条目丢失、其余元数据或版本漂移、集合身份变化、计划/文档改变时重新规划，不能覆盖当前修改。最终仍用原有 `verify/finalize` 检查集合与引用、总库保留以及真实原生证据。

Local API 执行器每次写前再次读取真实记录/版本与当前批次状态；成员更新只改该条目的 collections，保留其余成员关系，不发送删除条目请求。导入严格复用核实的身份与元数据，并按接口真实上限分块。写前持久记录开始状态，写后取得实际 URI/版本并重新读取核实；进程中断留下的 STARTED 或超时未知不能直接重放。

## 状态与恢复

`batch-status` 需要当前保存稿哈希和绑定快照。若提供观察回执，schema 为 `zotero-word.batch-observations.v1`，含原计划 `plan_fingerprint` 和 `results`。每项记录：

- `operation_id`、`outcome`、实际 `observed_at`、具体 `evidence`，成功导入可记录实际 `uri`。
- `outcome` 可为 `succeeded`、`failed`、`unknown`、`not_applied`。同一状态回执每个 operation 仅保留一项最新已核实结果，原始工具日志另存。
- 写入不明确时保持 `VERIFY_OUTCOME`，不进入可重试批次。只有实际确认未写入、记录 `not_applied` 和 `no_write_verified: true`，且取得更新后的完整观察，才允许重试。
- 实际条目已存在或目标成员状态已实现时，重查标为 `VERIFIED_PRESENT` / `VERIFIED_APPLIED`，不重复操作。新出现多个相同标识的条目仍需身份确认。

成功文字不能代替当前库状态。库快照必须晚于回执中的完成/未知写入观察，集合快照必须晚于成员写入观察；只读公开检索不要求逐查询重新导出 Zotero。不得刷新时间标签冒充新观察。回执与计划的指纹不符、未经上述预期成员更新解释的库记录/版本漂移均拒绝继续旧批次。`OBSERVED_BATCH_EFFECTS` 只表示本批效果已在快照中核实，不是稿件完成、Word 更新或重开关联的 PASS。

完整计划和回执留文件，模型优先读取摘要、待确认项和下一批；不为每条参考文献回传整库或重跑全稿回归。日常流程与开发回归仍按 [长稿执行](large-manuscripts.md) 区分。

运行控制的文档操作锁与执行器的批次写入日志用途不同：前者阻止同稿并发原生操作，后者防止不明写入重复派发。任一记录都不是成功凭据；最终通过仍需真实新鲜库/集合状态和保存稿验收。纯解析 memo 不能缓存外部库观察或写入结果。
