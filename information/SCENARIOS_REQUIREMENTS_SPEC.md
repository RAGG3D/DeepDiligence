# Scenarios Generation - 用户需求规范
**文档版本**: 1.0
**创建日期**: 2026-02-28
**适用于**: Gemini Deep Research + Scenarios Sheet生成完整workflow

---

## 用户核心要求总览

本文档整理了用户对Scenarios生成workflow的所有明确要求。所有要求均为**强制性**，必须在每次研究中严格遵守。

---

## 第一部分：Indication定义的严格要求

### ⚠️ 要求1: **绝对禁止模糊定义**

**用户原话**:
> "我最后强调一次，不允许写Multiple/Solid tumor这种模糊定义，必须严格查阅全网资料，找到确切的已公布可查的indication范围"

**具体要求**:
1. **禁止使用**以下模糊术语:
   - ❌ "Multiple indications"
   - ❌ "Solid tumor"
   - ❌ "Advanced malignancies"
   - ❌ "Various cancers"
   - ❌ 任何不具体的癌症描述

2. **必须使用**明确的癌症类型:
   - ✅ "Non-Small Cell Lung Cancer (NSCLC)"
   - ✅ "Triple Negative Breast Cancer (TNBC)"
   - ✅ "Biliary Tract Cancer (BTC)"
   - ✅ "Colorectal Cancer (CRC)"
   - ✅ "Small Cell Lung Cancer (SCLC)"

3. **查证要求**:
   - 必须查阅**全网资料**（公司10-K、press releases、ClinicalTrials.gov、investor presentations、conference transcripts）
   - 找到**确切的、已公布可查的**indication范围
   - 每个indication必须有**明确来源引用**（NCT编号、10-K页码、press release日期）

4. **如果资料不明确**:
   - 不得猜测或推断
   - 必须在报告中明确说明"未找到具体indication信息"
   - 不得使用模糊术语填补空白

**示例对比**:

| ❌ 错误示例 | ✅ 正确示例 |
|------------|------------|
| CTX-471 (CD137, Multiple) | CTX-471 (CD137, NSCLC/SCLC/Melanoma/HNSCC/Mesothelioma) |
| CTX-10726 (FOLR1, Solid Tumors) | CTX-10726 (FOLR1, GC/HCC/Endometrial/RCC) |
| Drug X for Advanced Malignancies | Drug X (Target, [list specific cancer types from trial]) |

**实施方法**:
- Gemini prompt必须明确要求："Do NOT use vague terms like 'Multiple' or 'Solid tumor'. List all specific cancer types mentioned in clinical trials, 10-K, or press releases."
- generate_scenarios.py parser必须检测并拒绝模糊术语
- 如发现"Multiple"/"Solid"等词，报错并要求用户提供具体cancer types

---

## 第二部分：Market Share分析的粒度要求

### 要求2: **每个不同癌症必须单独分析**

**用户原话**:
> "如果确认是不同的癌症，则要Gemini deep research对每一种癌症数据进行市场份额分析预测，比如CTX-8371要对NSCLC和TNBC分别做分析"

**具体要求**:
1. **默认规则**: 每个不同的cancer type必须**单独**进行:
   - TAM (Total Addressable Market) 分析
   - Competitive landscape分析
   - Differentiation assessment
   - Market share projection (2024-2038逐年预测)

2. **单独分析的强制内容**:

   对于**每个cancer type**，Gemini必须提供:

   a) **TAM Analysis**:
   ```
   - 2024: XX,XXX patients ([specific line of therapy], [biomarker status if relevant])
   - Growth rate: X% CAGR
   - 2030: XX,XXX patients
   - 2038: XX,XXX patients
   ```

   b) **Competitive Landscape** (该cancer type的竞品):
   ```
   | Drug | Revenue (2024) | ORR | PFS | OS | Safety | Route |
   ```

   c) **Differentiation** (与该cancer type竞品的对比):
   ```
   - ORR: X% vs competitor Y%
   - PFS: X months vs Y months
   - Line-matched comparison (2L vs 2L, NOT 2L vs 1L)
   - Assessment: Best-in-class / Above-average / Average
   ```

   d) **Market Share Projection** (该cancer type的逐年份额):
   ```
   | Year | Market Share |
   | 2024 | 0% |
   | 2025 | 0% |
   ...
   | 2038 | X% |
   ```

3. **示例要求**:

   **CTX-8371 (PD-1xPD-L1, NSCLC/TNBC/HL/HNSCC/Melanoma)** 必须分为:

   - **Indication 1: NSCLC, 2nd-line**
     - TAM: [NSCLC specific patient numbers]
     - Competitors: [NSCLC 2L drugs with data]
     - Differentiation: [vs NSCLC 2L competitors]
     - Market Share: [NSCLC 2024-2038]

   - **Indication 2: TNBC, 2nd-line**
     - TAM: [TNBC specific patient numbers]
     - Competitors: [TNBC 2L drugs with data]
     - Differentiation: [vs TNBC 2L competitors]
     - Market Share: [TNBC 2024-2038]

   - **Indication 3: Hodgkin Lymphoma**
     - TAM: [HL specific]
     - Competitors: [HL drugs]
     - Differentiation: [vs HL competitors]
     - Market Share: [HL 2024-2038]

   - ... (依此类推，每个cancer type单独分析)

---

### 要求3: **混合Market Share的特殊情况**

**用户原话**:
> "除非有明确的证据表明市场把多个indication的份额混在一起price in（依旧参考DCF Template.xlsx中的BT7480，mUC单独forecast market share，其余癌症混合一起)"

**用户补充要求 (2026-02-28)**:
> "如果某药物在多个癌症上前景仍然模糊（未进入trial/社交媒体无主流看法/暂无医疗数据披露等等），但仍包含已经priced in的价值，那么允许合并，但deep research需要先逐个排查，如果有前景明确的癌症则分离出来单独进行forecast"

**具体要求**:

0. **必须先逐个排查** (MANDATORY FIRST STEP):
   - 对每个indication逐一检查: 是否进入trial？社交媒体有无主流看法？公司有无披露医疗数据？
   - **前景明确的** → 必须分离出来单独forecast
   - **前景模糊的** → 允许合并，但必须列出具体癌症名称

1. **混合处理的条件** (针对**前景模糊**的indications):

   a) **前景模糊的定义** (满足以下大部分条件):
   - 未进入clinical trial（无active NCT编号）
   - 社交媒体/分析师无主流看法（<10% analyst mentions）
   - 公司暂无披露该indication的医疗数据
   - 但仍可能被市场部分priced in（data transfer等）

   b) **命名规则** (STRICT):
   - ✅ "NSCLC/SCLC/Mesothelioma/HNSCC Combined"
   - ✅ "HNSCC/NSCLC/OV/BRCA/GC/ESCA Combined"
   - ❌ "Other Solid Tumors Combined" (FORBIDDEN)
   - ❌ "Other Cancers" (FORBIDDEN)
   - ❌ 任何使用"Solid Tumor/Solid Tumors"的名称 (FORBIDDEN)

2. **DCF Template参考案例**:

   **BT7480 (Nectin-4/CD137, mUC/HNSCC/NSCLC/OV/BRCA/GC/ESCA)**:
   - Row 17: Asset name列出所有indications
   - Row 18: **mUC Market Share** (单独预测) — 因为mUC是primary indication，analyst重点关注
   - Row 19: **Other Cancers Market Share** (其余混合) — HNSCC/NSCLC/OV/BRCA/GC/ESCA作为secondary indications混合处理

   **判断依据**:
   - mUC: 60%+ analyst mentions, company focus, larger TAM → 单独预测
   - Others: 各自<10% mentions, 合计40%, smaller individual TAMs → 混合处理

3. **如何让Gemini判断**:

   Gemini prompt必须包含:
   ```
   For each asset with multiple indications, analyze:

   1. Market Evidence:
      - What % of analyst reports mention each indication?
      - Does X/Twitter separate indications or discuss as a group?
      - Does company provide separate revenue guidance?

   2. Indication Hierarchy:
      - Is there a PRIMARY indication (>50% analyst focus)?
      - Are there SECONDARY indications (<10% each mention)?

   3. Decision Rules:
      - If ONE indication >50% analyst focus → separate MS row for primary + "Other Cancers" combined row
      - If NO clear dominant indication AND indications discussed as group → single combined MS row "All Indications"
      - If EACH indication discussed separately (>20% each) → separate MS row for EACH

   4. Output Format:
      - PRIMARY indication: "[Drug] [Indication] Market Share"
      - SECONDARY combined: "[Drug] Other Cancers Market Share"
      - ALL combined: "[Drug] Market Share" (only if evidence shows no separation)
   ```

4. **默认策略**:
   - **When in doubt, separate** — 如果不确定是否应该混合，默认为每个indication单独分析
   - 混合处理是**例外**，不是常规

---

### 要求4: **让Gemini自行判断划分策略**

**用户原话**:
> "让Gemini deep research自行判断如何划分市场，然后分别预测每个市场份额"

**具体要求**:
1. **Gemini的判断职责**:
   - 分析每个asset的所有indications
   - 根据要求3的决策规则，判断应该:
     - 全部单独预测 (default)
     - Primary单独 + Secondary混合 (BT7480模式)
     - 全部混合 (rare, 需要强有力证据)

2. **Gemini必须在报告中明确说明**:
   ```markdown
   **Market Share Forecasting Strategy:**

   Decision: [Separate all / Primary separate + Secondary combined / All combined]

   Rationale:
   - Analyst coverage breakdown: [X% indication A, Y% indication B, ...]
   - Social sentiment: [Separated / Grouped discussion]
   - Company guidance: [Separate / Combined revenue forecasts]
   - TAM comparison: [Indication A: XX,XXX pts, Indication B: YY,YYY pts, ...]

   Conclusion: [Explain why this strategy was chosen]
   ```

3. **Output Format要求**:

   根据Gemini判断，输出相应格式:

   **Strategy A: 全部单独** (default)
   ```markdown
   **Indication 1: NSCLC, 2nd-line**
   - Market Share Projection (NSCLC): [table]

   **Indication 2: TNBC, 2nd-line**
   - Market Share Projection (TNBC): [table]
   ```

   **Strategy B: Primary单独 + Secondary混合** (BT7480模式)
   ```markdown
   **Indication 1: mUC, 2nd-line** (Primary - 60% analyst focus)
   - Market Share Projection (mUC): [table]

   **Other Indications: HNSCC/NSCLC/OV/BRCA/GC/ESCA** (Secondary - combined 40%)
   - Market Share Projection (Other Cancers): [table]
   ```

   **Strategy C: 全部混合** (rare, 需要强证据)
   ```markdown
   **All Indications: [list]** (Market treats as single opportunity)
   - Evidence: [strong evidence that market prices all together]
   - Market Share Projection (All Indications): [table]
   ```

---

## 第三部分：Data Transfer分析要求

### 要求5: **识别非trial indications的priced-in情况**

**用户原话** (来自之前的要求):
> "一个治疗A/B/C癌症药物只有癌症A的trial，但市场认为根据药理病理A的数据可以transfer到癌症B因此也priced in了癌症B，那么该药物的Scenarios就要计入癌症A和B的market shares"

**具体要求**:
1. **Data Transfer的定义**:
   - Drug有A/B/C三个potential indications
   - 只有indication A有active clinical trial
   - 但市场认为A的data可以transfer到B/C
   - 因此B/C虽然没有trial，但也被部分priced in

2. **如何识别Data Transfer**:

   a) **Pathology分析**:
   - 疾病生物学相似性 (如neuroendocrine features)
   - 分子标志物overlap (如DLL3高表达)

   b) **Pharmacology分析**:
   - Target expression在不同cancer types的prevalence
   - Mechanism of action的适用性

   c) **Market Evidence**:
   - Analyst reports提到non-trial indications的% (即使<10%也要注意)
   - X/Twitter讨论中是否提及potential expansion
   - 公司10-K是否列为"potential future indication"
   - Conference Q&A是否有相关问题

3. **Data Transfer的处理**:

   如果判断某个non-trial indication被**partially priced in**:
   - 必须单独预测该indication的market share
   - Peak market share应**显著低于**primary indication (通常30-50% of primary)
   - Launch timing应**晚于**primary indication (通常延后2-3年)
   - 在报告中明确标注为"Data Transfer Indication"

4. **示例** (CTX-009):
   ```markdown
   **Asset: CTX-009 (DLL3, BTC/CRC/SCLC)**

   **Indication 1: Biliary Tract Cancer** (Primary - Active Trial NCT05506943)
   - Market Share: Peak 28% in 2030

   **Indication 2: Colorectal Cancer** (Trial NCT05513742 Completed)
   - Market Share: Peak 22% in 2030

   **Indication 3: Small Cell Lung Cancer** (Data Transfer - No Active Trial)
   - Data Transfer Analysis:
     - Pharmacology: DLL3 expression 80%+ in SCLC (vs 60% in BTC)
     - Pathology: Neuroendocrine features similar
     - Market Evidence: 15% analyst mentions, company 10-K lists as potential
     - Conclusion: PARTIALLY priced in
   - Market Share: Peak 12% in 2032 (43% of BTC peak, 2-year delay)
   ```

---

## 第四部分：Competitive Benchmarking要求

### 要求6: **Line-Matched横向对比**

**用户原话** (来自之前的要求):
> "注意要考虑对比的合理性，比如要用2线药物和2线药物疗效比等等"

**具体要求**:
1. **Line-Matched原则**:
   - 2线药物 vs 2线竞品 (NOT vs 1线)
   - 3线药物 vs 3线竞品 (NOT vs 1/2线)
   - 1线药物 vs 1线竞品

2. **其他匹配原则**:
   - Biomarker-matched: PD-L1+ vs PD-L1+, EGFR+ vs EGFR+
   - Population-matched: Similar prior treatment history
   - Setting-matched: Metastatic vs metastatic, adjuvant vs adjuvant

3. **禁止的对比**:
   - ❌ 2线药物 vs 1线药物的ORR/PFS对比
   - ❌ 不同biomarker populations (PD-L1+ vs unselected)
   - ❌ 不同prior treatments (IO-naive vs IO-refractory)

4. **Competitive Data必须包含**:
   ```
   | Drug | Line | Revenue (2024) | ORR | PFS | OS | Grade 3+ AEs | Route |
   |------|------|----------------|-----|-----|----|--------------  |-------|
   ```

---

## 第五部分：Source Citation要求

### 要求7: **每个结论必须有明确来源**

**用户原话** (来自之前的要求):
> "同时一定让Gemini明确给出药物indication的来源"

**具体要求**:
1. **必须引用的来源类型**:
   - **NCT编号**: NCT05506943
   - **10-K页码**: "10-K FY2024 p.45"
   - **Press Release**: "Press release dated 2024-11-15"
   - **Analyst Report**: "Jefferies equity research 2024-11-15"
   - **Conference**: "ASCO 2024 oral presentation"
   - **ClinicalTrials.gov**: Full trial record link

2. **每个indication必须引用**:
   ```markdown
   **Indication: Biliary Tract Cancer, 2nd-line**
   - Source: NCT05506943 (Phase 2/3, active)
   - Additional: Company 10-K FY2024 p.45, Investor presentation Nov 2024
   ```

3. **Market evidence必须引用**:
   ```markdown
   - Analyst coverage: 60% (Jefferies 2024-11-15, SVB Leerink 2025-01-10, Canaccord 2024-12-05)
   - X/Twitter: ~65% of 150 mentions in past 6 months
   ```

4. **Competitive data必须引用**:
   ```markdown
   | Pemigatinib | $120M (Incyte Q3 2024 earnings) | 35% (FIGHT-202 trial) | ...
   ```

---

## 第六部分：Output Format要求

### 要求8: **Scenarios Sheet格式规范**

**基于用户要求和DCF Template 2020.xlsx**:

1. **Asset Name Row格式**:
   ```
   [Drug Name] ([Target], [Indication1]/[Indication2]/[Indication3]/...)

   例子:
   - CTX-009 (DLL3, BTC/CRC/SCLC)
   - BT5528 (EphA2, mUC/OV/NSCLC/HNSC/TNBC/GC)
   - Zelenectide (BT8009, Nectin-4, mUC)
   ```

2. **Market Share Row格式**:

   根据Gemini判断策略 (要求4):

   **Strategy A: 每个indication单独**
   ```
   [Drug Name] ([Target], [All Indications]) [Specific Indication] Market Share

   例子:
   - CTX-009 (DLL3, BTC/CRC/SCLC) BTC Market Share
   - CTX-009 (DLL3, BTC/CRC/SCLC) CRC Market Share
   - CTX-009 (DLL3, BTC/CRC/SCLC) SCLC Market Share
   ```

   **Strategy B: Primary单独 + Secondary混合**
   ```
   [Drug] ([Target], [All]) [Primary Indication] Market Share
   [Drug] ([Target], [All]) Other Cancers Market Share

   例子:
   - BT7480 (...) mUC Market Share
   - BT7480 (...) Other Cancers Market Share
   ```

   **Strategy C: 全部混合** (rare)
   ```
   [Drug Name] ([Target], [All Indications]) Market Share

   例子:
   - BT1718 (MT1, NSCLC/ESCA/NSCLSarcoma) Market Share
   ```

3. **禁止的格式**:
   - ❌ `CTX-471 (CD137, Multiple) Market Share`
   - ❌ `Drug X (Target, Solid Tumors) Market Share`
   - ❌ `Drug Y Market Share` (没有列出indications)

---

## 第七部分：Workflow执行checklist

### 执行每个ticker研究时，必须确认以下所有项:

**Phase 1: Clinical Trials数据获取**
- [ ] 运行 `clinical_trials_fetcher.py`
- [ ] 验证找到所有NCT编号
- [ ] 检查JSON输出包含完整conditions列表
- [ ] **Critical**: 确认没有"Advanced Malignancies"等模糊描述

**Phase 2: Gemini Deep Research**
- [ ] 加载clinical trials JSON数据到prompt
- [ ] **Critical Check 1**: Gemini输出的每个indication都是**具体癌症名称**（不允许Multiple/Solid tumor）
- [ ] **Critical Check 2**: 每个不同癌症都有**单独的TAM/Competitive/MS分析**
- [ ] 验证market share forecasting strategy有明确说明（单独 vs 混合）
- [ ] 验证每个indication有NCT来源引用
- [ ] 验证competitive data有line-matched对比
- [ ] 验证data transfer分析（如适用）

**Phase 3: Scenarios Sheet生成**
- [ ] Parser正确解析multi-indication格式
- [ ] 每个indication生成单独market share row（根据strategy）
- [ ] Asset名称列出所有具体cancer types
- [ ] 验证没有模糊术语（Multiple/Solid/Advanced）
- [ ] 在Excel中打开验证格式正确

**Phase 4: 质量验证**
- [ ] 随机抽取3个indications，验证NCT来源可查
- [ ] 随机抽取3个competitive drugs，验证revenue/ORR数据可查
- [ ] 检查是否有模糊术语漏网
- [ ] 验证market share strategy合理性

---

## 第八部分：Error Handling

### 如果遇到以下情况，必须采取的行动:

1. **发现模糊术语** (Multiple/Solid tumor):
   - 停止处理
   - 报错: "Vague indication term detected: [term]. Please specify exact cancer types."
   - 要求用户查证或提供具体cancer list

2. **无法找到具体indications**:
   - 在报告中明确说明: "No specific indication information found in available sources"
   - 列出已查证的来源: 10-K, press releases, ClinicalTrials.gov, etc.
   - 不得使用模糊术语填补空白

3. **Market share strategy不确定**:
   - Default to separate (每个indication单独分析)
   - 在报告中说明: "Due to insufficient market evidence, defaulting to separate forecasts for each indication"

---

## 附录A：DCF Template 2020.xlsx参考案例

### Case 1: Zelenectide (单一indication)
```
Row 10: Zelenectide (BT8009, Nectin-4, mUC)
Row 11: Zelenectide (BT8009, Nectin-4, mUC) Market Share
```
**Strategy**: 单一indication → 单一market share row

### Case 2: BT5528 (多个indications，主要关注一个)
```
Row 12: BT5528 (EphA2, mUC/OV/NSCLC/HNSC/TNBC/GC)
Row 13: BT5528 (EphA2, mUC/OV/NSCLC/HNSC/TNBC/GC) mUC Market Share
```
**Strategy**: mUC是primary indication → mUC单独预测（其余可能有combined row或不预测）

### Case 3: BT1718 (多个indications，作为整体)
```
Row 15: BT1718 (MT1, NSCLC/ESCA/NSCLSarcoma)
Row 16: BT1718 (MT1, NSCLC/ESCA/NSCLSarcoma) Market Share
```
**Strategy**: 市场作为整体定价 → 单一combined market share row

### Case 4: BT7480 (多个indications，primary单独+secondary混合)
```
Row 17: BT7480 (Nectin-4/CD137, mUC/HNSCC/NSCLC/OV/BRCA/GC/ESCA)
Row 18: BT7480 (Nectin-4/CD137, mUC/HNSCC/NSCLC/OV/BRCA/GC/ESCA) mUC Market Share
Row 19: (可能有) BT7480 (...) Other Cancers Market Share
```
**Strategy**: mUC dominant (60%+ analyst focus) → mUC单独，其余混合

---

## 附录B：用户原话汇总

为确保理解准确，以下是用户的关键原话：

1. **关于模糊定义**:
   > "我最后强调一次，不允许写Multiple/Solid tumor这种模糊定义，必须严格查阅全网资料，找到确切的已公布可查的indication范围"

2. **关于单独分析**:
   > "如果确认是不同的癌症，则要Gemini deep research对每一种癌症数据进行市场份额分析预测，比如CTX-8371要对NSCLC和TNBC分别做分析"

3. **关于混合情况**:
   > "除非有明确的证据表明市场把多个indication的份额混在一起price in（依旧参考DCF Template.xlsx中的BT7480，mUC单独forecast market share，其余癌症混合一起)"

4. **关于Gemini判断**:
   > "让Gemini deep research自行判断如何划分市场，然后分别预测每个市场份额"

5. **关于对比合理性**:
   > "注意要考虑对比的合理性，比如要用2线药物和2线药物疗效比等等"

6. **关于data transfer**:
   > "一个治疗A/B/C癌症药物只有癌症A的trial，但市场认为根据药理病理A的数据可以transfer到癌症B因此也priced in了癌症B，那么该药物的Scenarios就要计入癌症A和B的market shares"

7. **关于来源引用**:
   > "同时一定让Gemini明确给出药物indication的来源"

8. **关于标注具体癌症**:
   > "在Scenarios表格时一定要注明你提取的market share具体针对的是哪（几）个癌症"

9. **关于workflow确保**:
   > "把这些要求全部写入workflow，确保以后每次研究都按照所有我给你的要求来"

---

## 文档维护

- **版本**: 1.0
- **最后更新**: 2026-02-28
- **维护者**: 项目所有者要求
- **强制性**: 本文档所有要求为强制性，不得省略或简化

---

**END OF REQUIREMENTS SPECIFICATION**
