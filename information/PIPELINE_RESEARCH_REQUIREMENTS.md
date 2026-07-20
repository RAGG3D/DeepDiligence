# Pipeline研报要求清单

## 零、公司主体与数据血缘门槛（任何建模前必须完成）

1. 先建立 `{TICKER}_company_facts.json`：法定名称、交易代码/交易所、交易及报告币种、财年结日、会计准则、基本股数及其 as-of 日期、最新现金/有价证券/受限现金/债务拆分。
2. 必须查清代码沿用、改名、反向并购、壳公司、分拆、ADR 比率调整、拆股/反向拆股。`price_history_start_date` 必须是**当前经营主体**开始具有经济连续性的首个交易日；前身或壳公司历史不得进入当前公司的 Lifetime High/Low/分位数。
3. 历史价格须使用同一复权口径，并仅使用已完成交易日；禁止把盘中 `currentPrice` 与未收盘 OHLC 混在同一快照。每个市场字段必须记录 provider、session date、currency 与 adjustment basis。
4. 股数优先使用最新 SEC/监管文件封面页的 exact basic shares outstanding，并与市场数据供应商交叉核对；不得用加权平均股数、稀释股数或 fully diluted 股数冒充基本流通股。
5. Cash、restricted cash、current/non-current marketable securities、debt 必须分别提取。只有公司明确给出并可重算时才能把 cash + investments 称为 unrestricted liquidity；禁止把 liquidity 统称为 cash。
6. 来源优先级：SEC/公司正式披露 > FDA/ClinicalTrials.gov/原始学会资料 > 交易所/监管机构 > 权威二手资料 > 聚合器/搜索摘要。搜索摘要不能作为最终来源。
7. 每个数值必须保存 `value + unit + as_of/data_cut + URL + classification`；classification 只能是 Reported Fact、Company Estimate/Claim、Market Data、Analyst Assumption。
8. 必须用 numerator/denominator 重算百分比。若新闻稿、10-K/10-Q、演示材料或海报冲突，列出冲突并暂停自动写入该字段，不得静默挑选。
9. 安全性必须分别记录 any-grade、Grade 3+、serious AE、停药及样本量；“generally mild/moderate”不得改写成“全部为轻中度”。
10. 所有 Analyst Assumption（价格、份额、上市年、TAM 推演）必须与已报告事实分栏，并展示公式，不能伪装成公司指引或外部事实。

## 一、研报结构与输出

1. 每个药物（pipeline asset）独立生成一份完整研报，以Word文档（.docx）保存在公司文件夹的`pipeline_base4/`子目录中
2. 研报文件命名：`{TICKER}_{DrugName}_research_{timestamp}.docx`
3. 每份研报必须包含以下完整章节：
   - 药物概况（靶点、机制、适应症列表、当前临床阶段）
   - 各适应症市场分析（TAM、竞争格局、差异化评估）
   - 上市时间预测（Stage 1-5时间线）
   - 上市后逐年Market Share Forecast（2024-2038）
   - 数据来源汇总

## 二、适应症定义（严格禁止模糊）

4. 绝对禁止使用以下模糊术语：Multiple indications / Solid tumor(s) / Advanced malignancies / Various cancers / 任何不具体的癌症描述
5. 必须列出所有具体癌症类型的标准缩写：NSCLC、TNBC、BTC、CRC、SCLC、mUC、HNSCC、OV、HL、Melanoma等
6. 每个适应症必须注明来源：NCT编号、10-K页码、press release日期、conference名称
7. 若资料不明确，必须声明"未找到具体indication信息"，不得用模糊术语填补

## 三、逐适应症独立分析

8. 每种不同的癌症类型必须进行完整的独立分析，包括：
   - TAM（Total Addressable Market）：具体治疗线别+生物标志物亚群的患者数量（2024→2030→2038）及增长率
   - 竞争格局：该适应症+该治疗线别的**所有**竞争药物（见第五章）
   - 差异化评估：与竞品的疗效/安全性/给药方式量化对比
   - 逐年Market Share预测表（2024-2038）

## 四、Forecast逻辑与数据要求

9. 每个market share预测数字必须有详细的数据证据支撑，不允许出现"illustrative"或"qualitative projection"
10. 推理过程必须缜密，需明确说明：该数字基于哪些竞品的市场份额分布、该药物相对竞品的疗效差异、预期上市后的ramp-up速度（参考同类药物历史曲线）
11. Peak market share必须有竞品对标依据（如"类似疗效优势的X药物在同一适应症达到了Y%份额"）
12. 份额下降阶段需考虑：专利到期、后发竞品上市、biosimilar冲击等因素

## 五、竞争格局（最关键要求）

13. **必须覆盖全部竞品**：针对同一市场（同适应症+同治疗线别）的患者群体，所有已上市药物和已进入临床阶段的研发中药物都必须纳入分析
14. 已上市竞品必须包含：药物名、年销售额（2024）、ORR、PFS（月）、OS（月）、Grade 3+不良反应率、给药方式（口服/静脉）
15. 临床阶段竞品必须包含：药物名、开发公司、当前临床阶段、已披露的疗效数据、预计上市时间
16. **必须估计研发中competitor的预计上市时间**（基于当前临床阶段、同类药物开发历史、公司资源）
17. 如果某研发中竞品满足以下**全部**条件，可以忽略不纳入分析：
    - 临床数据显示竞争力明显较弱（ORR/PFS显著劣于标准治疗）
    - 且预计上市时间比本药晚4年以上
18. 被忽略的竞品需简要说明忽略理由（一句话）

## 六、治疗线别匹配（Line-Matched对比）

19. 2线药物只能与2线竞品对比（NOT vs 1线）
20. 3线药物只能与3线竞品对比
21. 生物标志物亚群必须匹配：PD-L1+ vs PD-L1+、EGFR+ vs EGFR+
22. 患者人群必须匹配：转移性 vs 转移性、辅助治疗 vs 辅助治疗

## 七、多适应症药物的Forecast策略

23. Gemini必须先**逐一排查**每个适应症的前景清晰度：是否有active trial？分析师有无单独讨论？公司有无披露数据？
24. 前景清晰的适应症（有trial、有数据、有分析师关注）→ 必须单独forecast
25. 前景模糊的适应症（无trial、无主流社交/分析师看法、无数据披露）但仍含priced-in价值 → 允许合并，但合并名称必须列出所有具体癌症名称（如"HNSCC/NSCLC/OV Combined"，禁止写"Other Solid Tumors"）
26. 默认策略为"全部单独"，合并是例外
27. 必须明确说明forecast策略选择的依据（分析师覆盖度、社交媒体讨论、公司指引）

## 八、Data Transfer分析

28. 若药物有A/B/C三个潜在适应症但仅A有active trial，需分析B/C是否因药理/病理相似性被市场部分priced in
29. Data Transfer的判断依据：靶点在不同癌症的表达率、疾病生物学相似性、分析师/社交媒体是否提及扩展潜力、公司10-K是否列为future indication
30. 被Data Transfer priced in的适应症：peak market share应显著低于primary indication（通常30-50%），上市时间应延后2-3年

## 九、上市时间预测

31. 每个药物必须给出Stage 1-5（Phase I→II→III→BLA→Approval）的已知或预测年份
32. 已完成的阶段必须注明来源（press release/ClinicalTrials.gov）
33. 未来阶段的预测必须基于：同适应症竞品的历史开发周期、近年（2023-2025）同类分子的获批速度、公司规模和资金储备
34. 必须给出预测的reasoning逻辑（如"同靶点竞品X从Phase II到获批用了Y年"）

## 十、TAM表格数据参考

35. Gemini在分析时可以且应当参考DCF Template中TAM Solid+MM和TAM Blood两个sheet的现有数据，但这些属于内部数据集，必须保留原始来源与核验日期，包括：
    - 已上市竞品的历史销售额（逐年数据）
    - 各适应症的患者人数（incidence/prevalence数据）
    - 市场份额分布（各竞品在该适应症的实际市场份额）
    - Maturity Parameter（药物上市后份额增长曲线参考：Best-In-Class / Tier One / Average三档）
    - Single Indication Maturity Parameter（NSCLC单一适应症特药的份额增长基准）
    - COGS参数（小分子37%、单抗45%，用于估算竞品利润率和定价策略）
36. 禁止默认认为内部 TAM 表优先。若内部数据有可追溯、更新且同口径的一级来源，才可采用；否则以最新一级来源为准，并更新内部数据中心。
37. 若 TAM 表与外部资料冲突，必须先比较口径、时间、币种、患者线别和来源等级，记录 reconciliation；无法消除的冲突不得自动进入模型。

## 十一、Reasoning穷尽参照物原则

38. 进行任何预测推理时，必须**穷尽所有合适的参照物**，不得只引用一个便利样本
39. 具体规则：
    - 若某项预测只存在唯一一个合适参照物（如同靶点竞品仅1个），可以仅参考该参照物
    - 若存在多个合适参照物，**必须全部列出**，不得选择性引用
    - 例："同靶点ADC竞品的Phase II→获批周期"如果有3个竞品分别用了4年、5年、6年，必须全部列出，不得只引"最快的4年"
40. 在多个参照物中选择最终预测值时，采用**多数原则+相近度原则**：
    - 优先采用参照物中出现频次最高的数值区间
    - 在频次相同时，优先采用与本药物情况最相近的参照物（按下方分场景排序）
41. 参照物的相近度排序**按预测场景分别指定**：
    - **预测上市时间/开发周期**：① 同公司规模 > ② 同分子类型 > ③ 同靶点同适应症 > ④ 同机制
    - **预测Market Share / Peak Share**：① 同靶点同适应症 > ② 同机制同适应症 > ③ 同分子类型 > ④ 同公司规模
    - **预测Ramp-up速度（上市后份额爬坡）**：① 同公司规模 > ② 同靶点同适应症 > ③ 同分子类型 > ④ 同机制
42. 必须在研报中展示完整的参照物列表和最终选择的reasoning过程

## 十二、来源引用

43. 每个结论必须有明确来源：NCT编号、10-K页码、press release日期、分析师报告名称+日期、conference名称+日期
44. 竞品数据必须标注来源：年销售额来自哪个公司的earnings、ORR/PFS来自哪个trial
45. TAM患者数量必须标注来源：流行病学报告、IQVIA数据、公司investor presentation

## 十三、质量标准

46. 不允许出现无数据支撑的定性判断（如"有望获得较高市场份额"）
47. 不允许使用"illustrative"、"qualitative"、"hypothetical"等修饰词
48. 所有数字必须可追溯到具体来源
49. 如某项数据确实无法获取，明确声明"数据不可得"并说明原因，不得编造

## 十四、Catalyst 强制要求

50. Catalyst universe 必须来自 Scenarios 中全部 drug × indication market-share 行；一个药物的不同 indication 必须分列，禁止用 “Other Pipeline” 聚合或截断。
51. 无 active catalyst 时使用同一 v7 全量中性框架；active catalyst 时必须按每个 relevant target 的 Conviction >=10% 过滤 outcome，并对所有剩余 outcome 做笛卡尔积。Catalyst 主表与 Scenarios 的 Catalyst Scenarios 必须使用完全相同的连续 scenario ID。实际 Catalyst 的 Excel Data Table 必须内嵌于 Catalyst B8:C末行（C8 显示 3%，B9 显示 Base），使用 Catalyst 本地桥接 input cells 驱动 VALUATION C3/C5；VALUATION O4:P200 必须为空，其他 Valuation sensitivity tables 不得删除。
52. 收到“ticker + catalyst run”后，先核验公司最新正式披露及近期官方学会通告，写明 drug/indication/trial/phase/预计数据；主表固定顺序为 Scenario、Base Case、Final Market Price、Upside、RJConv.、每个 relevant target 的 outcome 列、全部 target breakdown。RJConv. 指 Raw Joint Conviction；Base RJConv. 为100%，scenario RJConv. 为独立假设下所选 outcome conviction 原始乘积并按其降序排列。所有非相关 target 组保留且全部列可见，以灰色底色和灰色数据字体 mask。非相关 target 的 LOA 及 Catalyst Scenarios market share 保持 Absolute/base 不变；每个 target 的 scenario breakdown value 必须等于该 scenario Base Case price × 锁定的第 6 行 target 占比，Final Market Price 汇总全部 target 的 risk-adjusted Market Price。主表删除 Catalyst Target / Outcome / Scenario Summary；下方 Table 3 每个 target 使用独立 MS/LOA/Conv./spacer 四列组。
53. 收到“ticker + post-catalyst”后，必须先把完整 XLSX、Catalyst XML、全页公式/输入/缓存值/样式、实际价格反应和全网解读提交到持久化数据库；事务成功后才可清空手工输入并恢复颜色。
54. 估值只加入当年 `MAX(0, RCFS Ending Cash) / basic shares outstanding`，不得加入负现金，也不得与旧 valuation-date cash 重复计算。
55. 每个 Catalyst Scenarios 区块的 asset title 行 Y 单元格必须为空；每个 relevant target 的 market-share Y 公式必须以 Absolute Y 值为基准，按该组合 outcome 回查 Catalyst Table 3 的 MS change。Suspension 为零，其他 outcome 使用 `MAX(0, Absolute MS + change)`，且所有行列引用必须显式锁定。

## 十五、Historical Events 强制要求

56. 必须遍历公司官网归档/RSS/分页并收集显示四年内全部 press release；同日多条不得丢失。零条官网结果视为失败，不得用二手新闻替代后交付。
57. 临床数据 Event 必须以披露场合开头，并按 Phase / data type / N / ORR / CR / survival / safety 顺序极简记录确切数字。
58. 学术会议数据必须同时阅读公司 news release 与会议 abstract/poster/report；Event 单元格优先超链接会议来源，否则链接公司数据稿。abstract 不公开时必须明确记录 unavailable。
59. 每个新 ticker 在填入前清空 Catalyst 手工输入/事件元数据及 Historical Events 旧事件/旧超链接，并在交付前执行旧 ticker、旧药物、旧 URL 和结构一致性审计。

## 十六、历史临床 Catalyst 回测指令

60. 收到“ticker + test + event”（例如 `CMPX test ASCO2026`）时，仅回测已经完成的临床披露，并在现有模型新增 `Test-EVENT` tab；未来或无法核实的 event 必须 fail closed。
61. Test 的医学判断阶段只能检索 clinical data：公司临床数据稿、学会 abstract/poster、trial registry、同行评审论文、监管临床文件及竞品 clinical readout；所有价格/商业字段必须从医学判断上下文剔除。必须先锁定 MS/LOA/Conviction，之后才允许读取最早公开披露日前 7 个日历日窗口内的已完成交易日 raw close。盲测阶段始终禁用披露当日及之后的价格、价格反应、收益率、成交量、市值、beta 和分析师目标价；冻结后独立评分仅适用第74–77条的窄例外。
62. 模型必须结合实际披露、`datastore/dd.duckdb` 中过滤后的 clinical-only peer metrics 与全网 clinical competitor evidence，自行给出每个 event target 的 Increase/Remain/Decrease/Suspension Market Share Change、LOA Change、Conviction 及医学 rationale；MS/LOA 均为绝对百分点变化，Conviction 合计 100%。
63. 排除每个 target Conviction <10% 的 outcome，再对全部保留 outcome 做笛卡尔积。Test tab 中 event target 保留彩色并前置；全部非 event drug × indication 仍须可见，以灰底灰字 mask，禁止隐藏。
64. `Test-EVENT` 必须展示 actual readout、enrolled/evaluable denominator、data cutoff、疗效/持久性/安全性、line/biomarker/population/dose/follow-up 匹配竞品、cross-trial limitations、可点击 clinical sources 及披露前价格校准证据；主表须与 Catalyst 一致，包含 Scenario、原生 Data Table Base Case、Final Market Price、Upside、RJConv.、event outcome 与全 target breakdown。Base RJConv. 为100%，scenario RJConv. 为独立假设下所选 outcome conviction 原始乘积并按其降序排列。Base row 全部 breakdown Market Price 之和必须与披露前一周平均收盘价保持 `<0.5%` 相对差异；允许引用现有 DCF/VALUATION/operating-model 公式生成 scenario mechanics，外部工作簿公式仍禁用。
65. 写入前后必须执行 price-blind clinical guard 与独立 price-window guard，并验证 Table inputs 与 research JSON 完全一致、Conviction filter、scenario 组合数量/唯一性、全 target 可见灰色 mask、Base price reconciliation、盲测区域无披露当日/后续行情、无禁用公式/文本/URL 及 XLSX ZIP 完整性。
66. 每个 `Test-EVENT` 必须在 Scenarios 新增或替换 `Test Scenarios - EVENT` 模块；Test tab 与模块使用 workbook-global 唯一 scenario ID。活动 target 的 Y 公式按 Test 页 outcome 回查 clinical MS change 并从 Absolute Y 计算，Suspension 为零；非 event target 保持 Absolute Y，asset-title Y 为空。实际 Catalyst refresh 不得删除历史 Test 模块。除第61条严格限定的披露前 raw-close Base 校准外，禁止其他 observed security-market 输入或输出。
67. Catalyst/Test 的 Excel COM 生成只重算本次修改涉及的 Scenarios、Catalyst/Test 及原生 Data Table 范围，保存缓存后统一规范 calc state；禁止因无关旧外部依赖执行 whole-workbook `CalculateFullRebuild`。每个 Catalyst/Test Scenarios 标题后只能有一个真实 scenario ID，模板 divider 后的旧空白行必须先规范化，禁止遗留重复数字标题或截断最后一个 indication。

## 十七、所有 Research 的数据库常驻扫描

68. 任意 research 启动时必须并行扫描 `research_fact`、peer 数据和对应内部模块；将已有事实、缺口和冲突加入研究上下文。发现新信息后必须立即 upsert 到 durable JSON seed 与 DuckDB 正确模块，研究未结束前保持 rescan active。
69. 对库中药物的所有 in-scope 竞品，必须尽最大可能补齐 ORR、CR、PR、mPFS/PFS、mOS/OS、Safety overview、具名 side effects、每项 side-effect rate、dose/regimen、actual price 和 forecast price；同时记录已披露的 DoR、EFS、DFS、RFS 等时间结局。实际价与预测价分开存储；其他时间结局不能替代 mPFS/PFS 与 mOS/OS 各自明确的 reported/unavailable 记录。
70. 未披露字段必须以 `unavailable + source + reason` 存储，严禁编造；冲突以 `conflict` 保存，不得静默覆盖。每条事实保留 value、unit、population、dose、as-of/data-cut、URL、source kind 与 classification。
71. Test 医学判断只读取 clinical/dose/trial 字段；数据库仍可保存其他 workflow 获得的 commercial facts，但不得让其进入 Test 的医学解读。
72. 置顶结构文件 `00_WORKFLOW_STRUCTURE.md` 是 canonical map；以后每次修改 workflow 必须在同一变更中同步更新该文件，并同步相关 workflow 文档、skill、builder、validator 与回归审计。
73. Catalyst 与 Test 主表必须在 Upside 右侧加入 `RJConv.`（Raw Joint Conviction）；Base 为100%，scenario 为独立假设下各活动药物所选 outcome 的 Table-3 Conviction 原始乘积。保留所有 >=10% outcome 的笛卡尔积，不重新归一化，并按 RJConv. 由高到低排列，精确并列用稳定 outcome key 排序。
74. Blind Test 完成后必须先冻结并哈希 clinical artifact、blind prediction 与原 Test sheet，再启动隔离的第二 agent。第二 agent 不得修改原 MS/LOA/Conviction、scenario 或 breakdown，也不得把 post-release 信息返回盲测 agent。
75. 第二 agent 按首次公开披露时间取三个合格完成交易日：可靠的盘中/盘前披露可纳入当日收盘；盘后、非交易日或时间不明从下一交易日开始。仅列这三日 raw Close，写在 `D2:F2`；每个活动药物给1–10整数评分并写在其 final-LOA 同列第2行。
76. 同一三日窗口读取 regular-session 60m unadjusted High，找全部最高 RJConv. 并列 scenario。若三日峰值未达 blind Final Market Price，记录 `blind−peak` USD/share 与相对 peak 超出百分比；若达到，记录 `peak−blind` USD/share 与相对 peak 少盈利百分比。禁止使用第三个交易日之后行情。
77. Post-release Close/High 进入 `price` 模块，评分/触达测试进入 `backtest` 模块；不得进入 Test clinical-only context。保存后必须验证 blind digest 未变、原生 Data Table 存续、RJConv. 缓存与排序正确、无删除线/断链且 ZIP 完整。

## 十八、Event 全会临床公司筛选指令

78. 收到 `event + EVENT` 或 `EVENT + event` 时，先解析会议正式名称、届次、日期、abstract embargo 与 presentation schedule，再判断为已发生或未来 event。
79. 已发生 event 只保留在该 event 实际发布 quantitative human clinical data 的美股生物医药公司；未来 event 只保留有公司/学会/accepted abstract 正式证据明确将发布 quantitative human clinical data 的公司。仅参会、投资者会议、trial-in-progress、入组进度、preclinical 或非量化 corporate presentation 均排除。
80. “重要 clinical data”至少含一项 patient-level 指标：ORR、CR、PR、DCR/CBR、DoR、PFS/mPFS、OS/mOS、EFS/DFS/RFS、MRD、有效患者结局、AE/TEAE/TRAE/SAE、DLT、停药或治疗相关死亡。
81. 公司须在 cutoff 时于 Nasdaq、NYSE 或 NYSE American 上市并属于 biotech/biopharma；OTC、私营及仅非美上市公司排除。历史 event 使用当时 ticker。
82. 市值 `<USD 1,000M` 检验必须在隔离分支执行：已发生 event 使用最早临床披露前最后一个完成交易日，未来 event 使用筛选时最新完成交易日。临床分支只接收 `capital_under_1000m` 布尔值，不得接收、保存或输出数值市值、股价、OHLC、收益、价格反应、成交量、beta、目标价或估值信息。
83. Event 筛选仍须运行第68–71条数据库常驻扫描；commercial 信息可写入数据库正确模块，但不得进入 event blind artifact/context/output。
84. 保存 `artifacts/event_screens/{event_slug}.json` 并执行 price-blind validator。最终回复只允许按 ticker 字母排序的可点击 ticker list；无合格公司时只返回 `NONE`，不得展示市值、价格点评、公司摘要或 near-miss 名单。

具体执行顺序与命令见 `information/CATALYST_HISTORICAL_EVENTS_WORKFLOW.md`。
