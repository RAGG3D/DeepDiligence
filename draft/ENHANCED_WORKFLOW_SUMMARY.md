# Enhanced Scenarios Workflow - Implementation Summary

**Date**: 2026-02-28
**Status**: ✅ Clinical Trials + Gemini Complete | ⏳ Scenarios Parser Pending

---

## 完成的功能

### 1. ✅ ClinicalTrials.gov数据集成

**新脚本**: `clinical_trials_fetcher.py` (378行)

**功能**:
- 从10-K/8-K提取NCT编号 (正则 `NCT\d{8}`)
- 调用ClinicalTrials.gov API v2获取试验详情
- 按公司名搜索sponsor试验
- 输出JSON: trials分类by drug/indication/phase

**测试结果 (CMPX)**:
```
Total NCT IDs: 5
- NCT05506943: CTX-009 Phase 2/3, Biliary Tract Cancer (Active)
- NCT06150664: CTX-8371 Phase 1, NSCLC/TNBC/HL/HNSCC/Melanoma (Recruiting)
- NCT03881488: CTX-471 Phase 1, Multiple solid tumors (Completed)
- NCT05513742: CTX-009 Phase 2, Colorectal Cancer (Completed)

By Indication:
  - Biliary Tract Cancer: 1 trial
  - Non Small Cell Lung Cancer: 1 trial
  - Triple Negative Breast Cancer: 1 trial
  - Colorectal Cancer: 1 trial
  - (etc.)
```

**用法**:
```bash
python clinical_trials_fetcher.py \
    --ticker CMPX \
    --company-name "Compass Therapeutics" \
    --output-dir "/mnt/c/Users/yzsun/Desktop/DD/CMPX"
```

**输出**: `CMPX_clinical_trials_TIMESTAMP.json`

---

### 2. ✅ Gemini研报增强Prompt

**修改脚本**: `gemini_research.py`

**新增功能**:
- `format_clinical_trials_data(trial_data)` — 格式化JSON为可读prompt文本
- `run_deep_research()` 接受 `clinical_trials_data` 参数
- main() workflow: 先fetch trials → 格式化 → 传给Gemini

**增强的Prompt部分**:

#### a) **Priced-In Indications Analysis** (新增)
要求Gemini明确标注:
- 哪些indication有active trial (来源: NCT编号)
- 哪些indication通过data transfer被priced in
  - 病理学相似性 (pathology)
  - 药理学相似性 (pharmacology, target expression)
  - 市场证据 (analyst reports, X/Twitter讨论)
- 每个indication的来源引用 (NCT, 10-K页码, analyst report名称)

**示例输出格式**:
```markdown
**Asset: CTX-009 (DLL3, BTC/CRC/SCLC)**

**Priced-In Indications:**
- Biliary Tract Cancer (BTC), 2nd-line
  - Source: NCT05506943 (Phase 2/3, active), 10-K p.45
  - Market Evidence: 70% analyst reports (Jefferies, SVB)
  - Social: X/Twitter 60% BTC-focused

- Colorectal Cancer (CRC), 3rd-line
  - Source: NCT05513742 (Phase 2, completed), 10-K p.48
  - Market Evidence: 30% analyst mention

- Small Cell Lung Cancer (SCLC) - Data Transfer
  - Pharmacology: DLL3 target highly expressed in SCLC (80%+)
  - Pathology: Neuroendocrine features similar
  - Market Pricing: Partially priced in (15% analyst mentions, no active trial)
  - Conclusion: Include with reduced market share (10% peak vs BTC 25%)
```

#### b) **Competitive Marketed Drugs Data** (增强)
要求Gemini提供:
- **Revenue** (最新年度, 从TAM sheets/earnings)
- **ORR** (Objective Response Rate, %)
- **PFS** (Progression-Free Survival, months)
- **OS** (Overall Survival, months)
- **Safety** (Grade 3+ AEs %)
- **Route** (Oral vs IV vs SC)

**示例表格**:
```
Biliary Tract Cancer, 2nd-line competitors:

| Drug | Revenue (2024) | ORR | PFS | OS | G3+ AEs | Route |
|------|----------------|-----|-----|----|----|-------|
| Pemigatinib | $120M | 35% | 6.9mo | 17.5mo | 64% | Oral |
| Ivosidenib | $95M | 23% | 2.7mo | 10.8mo | 46% | Oral |
| Gem/Cis | N/A | 26% | 8.0mo | 11.7mo | 71% | IV |

CTX-009 + Paclitaxel (Phase 2):
- ORR: 34% (DLL3+)
- PFS: 5.5mo
- G3+ AEs: 58%
- Route: IV
```

#### c) **Horizontal Comparison** (新增要求)
- **Line-matched**: 2线vs 2线, **NOT** 2线vs 1线
- **Biomarker-matched**: PD-L1+ vs PD-L1+ (如相关)
- **Population-matched**: 相似prior treatment history

#### d) **Social/Analyst Sentiment** (新增)
- X/Twitter biotech community讨论 (哪些indication被提及)
- Analyst reports分析 (哪些indication是focus)
- % of mentions per indication

#### e) **Multi-Indication Output Format** (新增)
```
Asset: Drug (Target, IND1/IND2/IND3)

Market Share Rows:
- Drug (...) IND1 Market Share
- Drug (...) IND2 Market Share
- Drug (...) IND3 Market Share
```

---

### 3. ⏳ generate_scenarios.py Parser更新 (待完成)

**需要的修改**:
1. 解析multi-indication asset名称:
   - `BT5528 (EphA2, mUC/OV/NSCLC/HNSC/TNBC/GC)`
   - 当前: 只解析为single "All" indication

2. 解析indication-specific market share rows:
   - `BT5528 (...) mUC Market Share`
   - `BT5528 (...) Other Cancer Market Share`
   - 当前: 只有一个market share row

3. 生成多个market share rows per asset:
   - 每个priced-in indication一行
   - 使用inlineStr避免shared string corruption

4. XML格式:
   ```xml
   Row 10: <c r="C10" s="63" t="inlineStr"><is><t>BT5528 (EphA2, mUC/OV/NSCLC)</t></is></c>
   Row 11: <c r="C11" s="66" t="str"><f>=C10&amp;" mUC Market Share"</f><v>BT5528 (...) mUC Market Share</v></c>
   Row 12: <c r="C12" s="66" t="str"><f>=C10&amp;" Other Market Share"</f><v>BT5528 (...) Other Market Share</v></c>
   ```

---

## 完整Workflow示例

```bash
# CMPX (Compass Therapeutics) 完整流程

# Step 1: 获取clinical trials数据
python clinical_trials_fetcher.py \
    --ticker CMPX \
    --company-name "Compass Therapeutics" \
    --output-dir "/mnt/c/Users/yzsun/Desktop/DD/CMPX"

# 输出: CMPX_clinical_trials_20260228_013033.json
# 5个trials: CTX-009 (BTC, CRC), CTX-8371 (multi-tumor), CTX-471

# Step 2: 运行增强Gemini研报
python gemini_research.py \
    --ticker CMPX \
    --company-name "Compass Therapeutics" \
    --output-dir "/mnt/c/Users/yzsun/Desktop/DD/CMPX"

# 输出: CMPX_gemini_research_20260228_HHMMSS.docx
# 内容:
#   - Part 1: Market cap breakdown
#   - Part 2: Priced-in indications analysis
#       • CTX-009: BTC (primary, NCT05506943), CRC (NCT05513742), SCLC (data transfer)
#       • CTX-8371: NSCLC/TNBC/HL/etc. (NCT06150664)
#       • Competitive drug data with revenue/ORR/PFS/OS
#   - Part 3: Stage timelines

# Step 3: 生成Scenarios sheet (待parser更新后)
python generate_scenarios.py \
    --ticker CMPX \
    --research-file "/mnt/c/Users/yzsun/Desktop/DD/CMPX/CMPX_gemini_research_*.docx" \
    --company-name "Compass Therapeutics"

# 输出: DCF CMPX.xlsx Scenarios sheet
# 格式:
#   Row 10: CTX-009 (DLL3, BTC/CRC/SCLC) | Stage timeline
#   Row 11: CTX-009 (...) BTC Market Share | [%] | market share data
#   Row 12: CTX-009 (...) CRC Market Share | [%] | market share data
#   Row 13: CTX-009 (...) SCLC Market Share | [%] | market share data
```

---

## 关键改进

### ✅ 数据准确性
- ClinicalTrials.gov官方数据 → 防止Gemini幻觉trial status
- NCT编号 + 10-K page number → 可验证的来源引用

### ✅ Data Transfer Detection
- 识别没有active trial但被市场priced in的indications
- 病理学/药理学相似性分析
- Social/analyst sentiment验证

### ✅ 竞品Benchmarking
- 量化differentiation (ORR, PFS, OS对比)
- Line-matched比较 (2线vs 2线)
- Revenue数据 → 市场规模参考

### ✅ 透明度
- Scenarios sheet明确标注specific cancers
- 避免模糊的"All Indications"
- 每个indication单独market share row

---

## 文件清单

### 新增文件
```
/home/nazdaq_44sun/Investment/auto_dcf/
├── clinical_trials_fetcher.py (378 lines, NEW)
├── SCENARIOS_ENHANCED_WORKFLOW.md (full documentation)
└── ENHANCED_WORKFLOW_SUMMARY.md (this file)
```

### 修改文件
```
/home/nazdaq_44sun/Investment/auto_dcf/
└── gemini_research.py
    ├── format_clinical_trials_data() (NEW)
    ├── run_deep_research() (enhanced with clinical_trials_data param)
    ├── RESEARCH_PROMPT_TEMPLATE (enhanced sections)
    └── main() (fetch trials before Gemini)

/home/nazdaq_44sun/.claude/projects/.../memory/
└── MEMORY.md
    └── Scenarios Sheet Generation (ENHANCED 2026-02-28)
```

### 待修改文件
```
/home/nazdaq_44sun/Investment/auto_dcf/
└── generate_scenarios.py
    └── TODO: Parser update for multi-indication format
```

---

## 测试结果

### ✅ clinical_trials_fetcher.py
```
Tested: CMPX
Result: SUCCESS
- Found 5 NCT IDs (3 from 8-K, 5 from sponsor search)
- JSON output well-formatted
- API calls successful (ClinicalTrials.gov v2)
```

### ✅ gemini_research.py
```
Modified: Prompt enhanced, clinical trials integration
Status: READY for testing
Next: Run full Gemini research with CMPX trials data
```

### ⏳ generate_scenarios.py
```
Status: Parser update pending
Requirement: Handle multi-indication format
  - Parse: "Drug (Target, IND1/IND2/IND3)"
  - Parse: "Drug (...) IND1 Market Share"
  - Generate: Multiple MS rows per asset
```

---

## 下一步行动

### 1. 测试完整workflow (CMPX)
```bash
# Step 1: 已测试 ✅
python clinical_trials_fetcher.py --ticker CMPX --company-name "Compass Therapeutics"

# Step 2: 待测试
python gemini_research.py --ticker CMPX --company-name "Compass Therapeutics"
# 检查Word文档: priced-in indications analysis是否完整
```

### 2. 更新generate_scenarios.py parser
- 修改PipelineAsset.market_shares结构: `{indication: {year: %}}`
- 修改regex parsing: 支持indication-specific MS rows
- 修改XML generation: 每个indication生成单独MS row

### 3. 完整端到端测试
- CMPX: CTX-009 (BTC/CRC/SCLC), CTX-8371 (NSCLC/TNBC/etc.)
- 验证Scenarios sheet格式匹配DCF Template 2020.xlsx

---

**生成时间**: 2026-02-28
**文档**: SCENARIOS_ENHANCED_WORKFLOW.md (详细), ENHANCED_WORKFLOW_SUMMARY.md (本文件)
**状态**: ✅ 2/3 steps complete, ⏳ parser update pending
