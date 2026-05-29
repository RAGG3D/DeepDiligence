# DCF Auto-Fill 用户手册

## 1. 内置文件路径

### 1.1 读取和保存路径

| 用途 | 路径 | 说明 |
|------|------|------|
| Excel 默认路径 | `/mnt/c/Users/yzsun/Desktop/DD/{TICKER}/DCF {TICKER}.xlsx` | 所有脚本共用此默认路径 |
| 通知脚本 | `/home/nazdaq_44sun/Investment/auto_dcf/notify.py` | 发送邮件到 yzsun0123@gmail.com |
| 备份文件 | `DCF {TICKER}_backup_{timestamp}.xlsx` | `excel_writer.py` 写入前自动创建，与原文件同目录 |
| DCF 模板基准文件 | `DCF Template 2020.xlsx` | `create_template.py` 复制 TAM/Peer Views 等 sheet |
| Pipeline 报告 | `DD/{TICKER}/pipeline_base4/` | Gemini Deep Research base case 报告 |
| Bull 报告 | `DD/{TICKER}/pipeline_bull2/` | Gemini bull case 数据收集报告 |
| Bear 报告 | `DD/{TICKER}/pipeline_bear3/` | Gemini bear case 数据收集报告 |
| Catalyst 报告 | `DD/{TICKER}/pipeline_catalyst/` | Gemini catalyst 数据收集报告 |
| Clinical Trials | `DD/{TICKER}/{TICKER}_clinical_trials_*.json` | ClinicalTrials.gov API 数据 |

所有脚本均支持 `--path` 参数覆盖默认路径。

### 1.2 外部数据源

| 数据源 | URL/API | 用途 | 使用脚本 |
|--------|---------|------|----------|
| SEC EDGAR XBRL | `data.sec.gov/api/xbrl/companyfacts/` | 财务数据 | sec_fetcher.py |
| SEC EDGAR Submissions | `data.sec.gov/submissions/` | 10-K 文件列表 | sec_fetcher.py |
| SEC EDGAR EFTS | `efts.sec.gov/LATEST/search-index` | 会议相关 8-K | fill_events.py |
| Yahoo Finance | `yfinance.download(ticker)` | 历史股价 | fill_events.py |
| GlobeNewswire | RSS feed | 新闻稿 | fill_events.py |
| Google News | RSS feed | 新闻和会议信息 | fill_events.py |
| ClinicalTrials.gov | `clinicaltrials.gov/api/v2/studies/` | 临床试验数据 | clinical_trials_fetcher.py |
| Gemini Deep Research | `deep-research-pro-preview-12-2025` | 药物市场研究 | gemini_research.py, gemini_research_4pass_demo.py |
| Gemini 2.5 Flash | `gemini-2.5-flash` | 分析推理 | gemini_research_4pass_demo.py |
| Anthropic Claude | Claude Haiku | 新闻摘要生成 | fill_events.py |

---

## 2. 脚本总览

```
┌──────────────────────────────────────────────────────────────────────┐
│                        DCF 完整工作流                                │
├──────────────────────────────────────────────────────────────────────┤
│                                                                      │
│  Step 1: create_template.py  ──────> DCF {TICKER}.xlsx (空白模板)     │
│                                                                      │
│  Step 2: main.py + sec_fetcher.py ─> 填充 FY DATA (XBRL 财务数据)    │
│          + excel_writer.py                                           │
│                                                                      │
│  Step 3: fill_events.py ───────────> 填充 Historical Events          │
│                                      (股价 + 新闻 + AI 摘要)         │
│                                                                      │
│  Step 4: fill_tam.py ──────────────> 填充 TAM Solid+MM / TAM Blood   │
│                                      (药物销售数据)                   │
│                                                                      │
│  Step 5: clinical_trials_fetcher.py > 获取临床试验数据 (JSON)         │
│                                                                      │
│  Step 6: gemini_research_4pass_demo.py                               │
│          Phase 1: Deep Research ───> 全网数据收集 (1 次搜索)          │
│          Phase 2: 4 次独立思考 ────> base/bull/bear/catalyst 报告     │
│                                                                      │
│  Step 7: generate_scenarios.py ────> 填充 Scenarios Sheet             │
│                                      (Absolute/Base/Bull/Bear/       │
│                                       Breakdown/Catalyst)            │
└──────────────────────────────────────────────────────────────────────┘
```

### 2.1 活跃脚本一览

| 脚本 | 用途 | 主要输入 | 主要输出 |
|------|------|----------|----------|
| `create_template.py` | 生成空白 DCF 模板 | ticker, unit | DCF {TICKER}.xlsx |
| `main.py` | CLI 入口，填充 FY DATA | ticker, years | DCF xlsx FY DATA sheet |
| `sec_fetcher.py` | SEC EDGAR 数据获取 | CIK | financial_data dict |
| `excel_writer.py` | 手术式 xlsx 修补 | patches dict | 修改后的 xlsx |
| `fill_events.py` | 填充 Historical Events | ticker | DCF xlsx HE sheet |
| `fill_tam.py` | 填充 TAM 药物数据 | xlsx path | TAM sheets |
| `clinical_trials_fetcher.py` | ClinicalTrials.gov 数据 | ticker, company | JSON 文件 |
| `gemini_research.py` | Gemini 单次研究 (base) | ticker, company | .md + .docx |
| `gemini_research_4pass_demo.py` | 4-pass: 1搜索→4报告 | ticker, drug | 4 组 .md + .docx |
| `generate_scenarios.py` | 生成 Scenarios Sheet | report files | DCF xlsx Scenarios sheet |
| `notify.py` | 邮件通知 | subject, body | Email |

---

## 3. 建立 Excel 文件：`create_template.py`

### 3.1 用法

```bash
# 基本用法（K USD 模式，自动检测 FYE 月份）
python create_template.py --ticker BHVN

# 指定 CIK（当 ticker 解析不正确时）
python create_template.py --ticker BHVN --cik 0001935979

# MM USD 模式（公司以百万为单位报告）
python create_template.py --ticker LNTH --unit MM

# 仅生成 FY DATA sheet（用于插入已有工作簿）
python create_template.py --ticker CMPX --fydata-only

# 非 USD 货币（IFRS 公司）
python create_template.py --ticker MOLN --currency CHF

# 完整参数
python create_template.py --ticker BHVN --base-year 2020 --he-years 2022 2023 2024 2025 \
    --fye-month 12 --unit K --cik 0001935979 --path "/custom/path.xlsx"
```

### 3.2 K USD 模式生成的 Sheet（默认）

生成 **3 个 Sheet**：`FY DATA` | `FY DATA K USD` | `Historical Events`

### 3.3 MM USD 模式生成的 Sheet

生成 **2 个 Sheet**：`FY DATA` | `Historical Events`

---

## 4. 各 Sheet 生成代码详解

### 4.1 FY DATA K USD — `_build_kusd_sheet()`

核心数据 Sheet，所有财务数据以千美元 (K USD) 为单位存储。

**Sheet 结构（121 行）：**

| 行号 | 内容 | 生成函数 |
|------|------|----------|
| R1-R4 | 标题行 — Ticker, FYE 月份, 年份列 (F-K) | 直接写入 |
| **R5-R24** | **Income Statement** | |
| R8 | Revenue | `_DR()` |
| R9-R10 | R&D / G&A | `_DR()` |
| R11 | Total OpEx = R9+R10 | `_SM()` |
| R12 | Loss From Ops = R8-R11 | `_FR()` |
| R14-R15 | Interest Income/Expense | `_DR()` |
| R18 | Net Income = R12+R16-R17 | `_FR()` |
| R19 | EPS = R18*1000/R20 | `_FR()` |
| R20 | Shares Outstanding | `_DR()` |
| **R26-R49** | **Income Statement Notes** | |
| R28-R36 | R&D Item 1-9 (ISN) | `_DR()` |
| R37 | Total R&D = SUM(R28:R36) | `_SM()` |
| R39 | Check R&D = R37 - R9 | `_CK()` |
| R42-R46 | G&A Item 1-5 (ISN) | `_DR()` |
| R47 | Total G&A = SUM(R42:R46) | `_SM()` |
| R49 | Check G&A = R47 - R10 | `_CK()` |
| **R51-R88** | **Balance Sheet** | |
| R53-R56 | Current Assets | `_DR()` |
| R60-R62 | Non-Current Assets | `_DR()` |
| R63 | Total Assets | `_FR()` |
| R67-R70 | Current Liabilities | `_DR()` |
| R74-R77 | Long-Term Liabilities | `_DR()` |
| R81-R84 | Equity | `_DR()` |
| R88 | Check BS = R63 - R86 | `_CK()` |
| **R90-R110** | **Balance Sheet Notes** | |
| R93-R97 | PP&E Items + Depreciation | `_DR()` |
| R100 | Check PP&E = R98 - R60 | `_CK()` |
| R103-R107 | Accrued Items (BSN) | `_DR()` |
| R110 | Check Accrued = R108 - R68 | `_CK()` |
| **R112-R121** | **Cash Flow Statement** | |
| R117-R120 | Operating/Investing/Financing/FX | `_DR()` |

### 4.2 FY DATA (SUMIFS) — `_build_mm_sheet()`

仅在 K USD 模式下生成。采用 **"Build then overlay"** 策略：
1. 先调用 `_build_kusd_sheet()` 生成完整布局
2. D 列 → 公式引用 K USD（`='FY DATA K USD'!D{r}`）
3. E 列 → `[MM USD]`
4. 数据单元格 → SUMIFS/1000 公式（**紫色字体** `#7030A0`）

### 4.3 Historical Events — `_build_he_sheet()`

4 个并排区块，每年 365 天行：Date | Share Price | DoD Chg | EVT | Category

---

## 5. 填充数据：`main.py` + `excel_writer.py`

### 5.1 用法

```bash
python main.py --ticker CMPX                                          # 基本
python main.py --ticker BHVN --years 2020 2021 2022 2023 2024 --cik 0001935979  # 指定年份+CIK
python main.py --ticker CMPX --dry-run                                 # 干跑
python main.py --ticker CMPX --unit K                                  # 强制单位
```

### 5.2 数据获取 (`sec_fetcher.py`)

`build_financial_data()` 返回 **5 元组**：
```python
(financial_data, rename_map, note_details, reporting_unit, currency)
```

### 5.3 Excel 写入 (`excel_writer.py`)

**写入策略：** 永远不调用 openpyxl `.save()`，使用 `zipfile` + `ElementTree` 手术式修补 XML。

```
1. _backup()                    → 创建时间戳备份
2. openpyxl.load_workbook() ×2  → 双工作簿读取 (data_only=False + True)
3. _collect_kusd_patches()      → 收集单元格写入计划
4. _apply_xlsx_patches()        → 手术式 XML 修补
```

---

## 6. 填充历史事件：`fill_events.py`

```bash
python fill_events.py BHVN
```

**6 个数据源：** Yahoo Finance News, GlobeNewswire RSS, SEC EDGAR 8-K, Google News RSS, SEC EDGAR EFTS (会议), Google News Conferences

**AI 摘要：** Claude Haiku，新闻 30 词 / 大波动分析 65 词

---

## 7. 填充 TAM 数据：`fill_tam.py`

```bash
python fill_tam.py                    # 填充 DCF Template 2020.xlsx
python fill_tam.py --path /custom/path.xlsx
```

填充 `TAM Solid+MM` 和 `TAM Blood` 两个 sheet 的 FY2024/FY2025 药物销售数据。
使用手术式 zip 修补（与 excel_writer.py 相同架构）。

---

## 8. 临床试验数据：`clinical_trials_fetcher.py`

```bash
python clinical_trials_fetcher.py --ticker CMPX --company-name "Compass Therapeutics"
```

- 从 10-K/8-K 提取 NCT 编号
- 查询 ClinicalTrials.gov API v2
- 按公司名搜索
- 输出 JSON: `{TICKER}_clinical_trials_{timestamp}.json`

---

## 9. Gemini Deep Research：4-Pass 架构

### 9.1 架构说明

```
Phase 1: Deep Research Agent (1 次全网搜索)
    ↓ 生成 30K+ chars 原始数据报告
Phase 2: 4 次独立 Gemini 2.5 Flash 调用
    ├── Base:     完整分析 + 市场份额预测 (BASE_REASONING_PROMPT)
    ├── Bull:     数据收集 — 最佳可比药物、加速审批先例 (BULL_DATA_PROMPT)
    ├── Bear:     数据收集 — 完整竞争格局、安全性、失败案例 (BEAR_DATA_PROMPT)
    └── Catalyst: 数据收集 — 催化剂日历、历史股价反应 (CATALYST_DATA_PROMPT)
```

### 9.2 用法

```bash
# 完整 4-pass（1 次搜索 + 4 次分析）
python gemini_research_4pass_demo.py --ticker CMPX --company-name "Compass Therapeutics" --drug CTX-009

# 跳过 Phase 1（复用已有数据报告）
python gemini_research_4pass_demo.py --ticker CMPX --company-name "Compass Therapeutics" --drug CTX-009 \
    --data-report /path/to/existing_data_report.md

# 只运行部分 scenario
python gemini_research_4pass_demo.py --ticker CMPX --company-name "Compass Therapeutics" --drug CTX-009 \
    --scenarios base bull

# 单次 base case 研究（旧版，单药物单报告）
python gemini_research.py --ticker CMPX --company-name "Compass Therapeutics"
```

### 9.3 Bull/Bear/Catalyst 数据收集格式

所有数据收集 prompt 使用 **Peer Views 转置表格格式**（一列一个药物 readout，行为数据字段）：

**Bull (5 张表)**:
1. 最佳可比药物 Readouts (Peer View 30+ 行)
2. Breakthrough/Accelerated Approval 先例
3. 竞争对手失败/延迟案例
4. Label Expansion 先例
5. 最佳 Launch 轨迹

**Bear (5 张表)**:
1. 完整竞争格局 — 全部药物 (Peer View 30+ 行)
2. 详细安全性对比 (Grade 3+ AE 逐项)
3. 令人失望的药物 Launch
4. 小公司商业化挑战
5. 支付方/报销数据

**Catalyst (6 张表)**:
1. 即将到来的催化剂日历
2. 历史数据 Readout + 股价反应 (Peer View)
3. 可比催化剂先例
4. 分析师共识
5. 会议日历 (未来 18 个月)
6. SOC 基准数据

---

## 10. 生成 Scenarios Sheet：`generate_scenarios.py`

```bash
# 从 base case 报告生成（单目录）
python generate_scenarios.py --ticker CMPX --report-dir /mnt/c/Users/yzsun/Desktop/DD/CMPX/pipeline_base4/

# 完整 3 模块（base + bull + bear）
python generate_scenarios.py --ticker CMPX \
    --report-dir /mnt/c/Users/yzsun/Desktop/DD/CMPX/pipeline_base4/ \
    --bull-dir /mnt/c/Users/yzsun/Desktop/DD/CMPX/pipeline_bull2/ \
    --bear-dir /mnt/c/Users/yzsun/Desktop/DD/CMPX/pipeline_bear3/
```

**生成 3 个模块：**

| 模块 | Scenario # | 内容 |
|------|-----------|------|
| Module 0 | Scenario 4 | Absolute Value — 所有药物完整数据 |
| Module 1 | Scenario 1/2/3 | Base / Bull / Bear — 引用 Absolute，不同 peak MS |
| Module 2 | Scenario 5+ | Break Down — 逐个药物累加 |
| Module 3 | Scenario 9+ | Catalyst — 每个催化剂正/负面场景 |

---

## 11. 5 个校验行

所有校验行应为 0，ABS > 0.1 时显示红色背景：

| 行 | 校验内容 | 公式 |
|----|----------|------|
| R39 | R&D Notes = IS R&D | `= R37 - R9` |
| R49 | G&A Notes = IS G&A | `= R47 - R10` |
| R88 | Total Assets = Liabilities + Equity | `= R63 - R86` |
| R100 | PP&E Notes = BS PP&E | `= R98 - R60` |
| R110 | Accrued Notes = BS Accrued | `= R108 - R68` |

---

## 12. 完整工作流示例

```bash
cd ~/Investment/auto_dcf

# Step 1: 生成空白模板
python create_template.py --ticker BHVN --cik 0001935979

# Step 2: 填充 FY DATA
python main.py --ticker BHVN --years 2020 2021 2022 2023 2024 --cik 0001935979

# Step 3: 填充 Historical Events
python fill_events.py BHVN

# Step 4: 获取临床试验数据
python clinical_trials_fetcher.py --ticker BHVN --company-name "Biohaven Ltd"

# Step 5: Gemini 4-pass 研究 (每个药物分别运行)
python gemini_research_4pass_demo.py --ticker BHVN --company-name "Biohaven Ltd" --drug BHV-7000

# Step 6: 生成 Scenarios Sheet
python generate_scenarios.py --ticker BHVN --report-dir /mnt/c/Users/yzsun/Desktop/DD/BHVN/pipeline_base4/
```

---

## 13. 环境变量

| 变量 | 用途 | 设置方式 |
|------|------|----------|
| `GEMINI_API_KEY` | Gemini API 密钥 | `export GEMINI_API_KEY='...'` |
| `ANTHROPIC_API_KEY` | Claude API 密钥 (fill_events.py) | `.env` 文件 |
| `GMAIL_APP_PASSWORD` | 邮件通知密码 | `~/.bashrc` |
