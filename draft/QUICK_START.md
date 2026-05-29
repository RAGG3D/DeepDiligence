# DCF Auto-Fill - Quick Start Guide

## 快速运行（CMPX 示例）

```bash
cd ~/Investment/auto_dcf

# Dry-run 模式（查看数据，不修改 Excel）
python main.py --ticker CMPX --dry-run

# 正式运行（写入 Excel）
python main.py --ticker CMPX

# 自定义年份
python main.py --ticker CMPX --years 2020 2021 2022 2023 2024

# 自定义 Excel 路径
python main.py --ticker BHVN --path "/custom/path/DCF BHVN.xlsx"
```

---

## ✅ 运行后验证清单

### 1. 检查日志输出

```bash
# 查看最近一次运行的关键信息
python main.py --ticker CMPX 2>&1 | grep -E "Backup|calcChain|K USD cells|Row not found|WARNING"
```

**期望看到**：
- ✅ `Backup created: ...` - 自动备份已创建
- ✅ `Removing xl/calcChain.xml to force formula recalculation` - 公式将重算
- ✅ `K USD cells written: 138` - 数据已写入（数量会因公司而异）

**⚠️ 需要注意的警告**：
- `Row not found in K USD sheet: ('BS', None, 'Marketable Securities')` → 模板缺少该行，需手动添加
- `Row not found in K USD sheet: ('BS', None, 'Operating Lease Liabilities, Current Portion')` → 同上

### 2. 打开 Excel 验证

```bash
# Windows (from WSL)
explorer.exe /mnt/c/Users/yzsun/Desktop/DD/CMPX/DCF\ CMPX.xlsx
```

**验证步骤**：

1. **打开 `FY DATA K USD` sheet**

2. **检查关键行（2020 年，F 列）**：
   - R53 (Cash): 应显示 **47,076**（不是 2）
   - R54 (AR): 应显示 **0**（不是 2）
   - R56 (R&D Incentives): 应显示 **0**（不是 2）
   - R62 (Other Assets): 应显示 **583**（包含 Restricted Cash）
   - R69 (Deferred Rev Current): 应显示 **0**（不是 2）
   - R77 (Other LT Liabilities): 应显示 **0**（不是 2）
   - R83 (AOCI): 应显示 **0**（不是 2）

3. **检查 Check 行（R88）**：
   - 公式：`=F63-F86`（总资产 - 总负债和权益）
   - **期望值**：**0**（所有年份：F, G, H, I, J 列）
   - 如果显示旧值（如 -10），按 **Ctrl+Alt+F9** 强制重算

4. **查看 2022-2024 年的大额数据**：
   - 如果模板有 "Marketable Securities" 行：
     - 2022: **151,663** K USD
     - 2023: **128,233** K USD
     - 2024: **83,239** K USD
   - 如果没有该行：在日志中会看到 "Row not found" 警告

---

## 🛠️ 常见问题修复

### Q1: Check 行不为 0

**原因**：公式未重新计算
**解决**：
1. 在 Excel 中按 **Ctrl+Alt+F9** 强制重算所有公式
2. 保存文件
3. 如果仍不为 0，检查是否有 "Row not found" 警告（缺少关键行）

### Q2: 日志显示 "Row not found: Marketable Securities"

**原因**：Excel 模板缺少该行
**影响**：2022-2024 年的短期投资数据（$83M-$151M）无法写入
**解决方案**（二选一）：

#### **方案 A：手动添加行到模板**
1. 打开 Excel
2. 在 `FY DATA K USD` sheet，找到 R54 (Accounts Receivable)
3. 在 R54 上方插入新行
4. 设置：
   - Col B: 留空
   - Col C: `BS`
   - Col D: `Marketable Securities`
   - Col F-K: 留空（脚本会填入）
5. 保存模板
6. 重新运行脚本

#### **方案 B：手动从日志复制数据**
1. 运行 dry-run 查看数据：
   ```bash
   python main.py --ticker CMPX --dry-run | grep "Marketable Securities"
   ```
2. 手动在 Excel 中创建该行并输入数值

### Q3: 某些年份的数据是 "–" (None)

**正常情况**：
- **AOCI 2020**: 公司首年报告，无累计其他综合收益
- **FX Effect on Cash**: US 公司通常无外汇影响
- **R&D Incentives Receivable**: US 公司无此项（UK 特有）

**异常情况**：
- 如果核心项目（如 Cash, R&D, G&A）显示 "–"，说明 XBRL 映射失败
- 检查公司的 CIK 是否正确
- 查看 SEC EDGAR 确认公司是否已提交相应年份的 10-K

---

## 📊 CMPX 测试基准（用于对比）

### 2020 年资产负债表（K USD）

| 项目 | 金额 |
|------|------|
| **Assets** | |
| Cash | 47,076 |
| Accounts Receivable | 0 |
| Prepaid Expenses | 3,126 |
| R&D Incentives Receivable | 0 |
| PP&E, Net | 1,126 |
| Operating Lease ROU | 0 |
| Other Assets | 583 (含 Restricted Cash $263K) |
| **Total Assets** | **51,911** |
| | |
| **Liabilities** | |
| Accounts Payable | 1,061 |
| Accrued Expenses | 1,571 |
| Deferred Revenue, Current | 0 |
| Debt, Current | 7,467 |
| Operating Lease Liab, Current | 0 |
| Long-Term Debt | 1,867 |
| Operating Lease Liab, LT | 0 |
| Deferred Revenue, LT | 0 |
| Other LT Liabilities | 0 |
| **Total Liabilities** | **11,966** |
| | |
| **Equity** | |
| Common Stock | 5 |
| Additional Paid-In Capital | 191,348 |
| AOCI | 0 |
| Accumulated Deficit | -151,408 |
| **Total Equity** | **39,945** |
| | |
| **Check (Assets - L - E)** | **0** ✅ |

---

## 🔧 高级用法

### 启用详细日志
```bash
python main.py --ticker CMPX --verbose
```

### 批量处理多个公司
```bash
#!/bin/bash
for ticker in CMPX BHVN RXRX; do
    echo "Processing $ticker..."
    python main.py --ticker $ticker
done
```

### 查看某个年份的详细 XBRL 数据
```bash
python -c "
from sec_fetcher import SECFetcher
fetcher = SECFetcher()
data = fetcher.build_financial_data('CMPX', [2024])
for key, vals in sorted(data.items()):
    print(f'{key}: {vals}')
"
```

---

## 📚 相关文档

- **`REFACTORING_SUMMARY.md`** - 详细的技术文档和修复说明
- **`MEMORY.md`** - 项目架构和注意事项（Claude AI 记忆文件）
- **`sec_fetcher.py`** - XBRL 概念映射表（查看所有支持的财务科目）

---

## 🆘 获取帮助

遇到问题时：

1. **检查日志**：运行 `--verbose` 模式查看详细输出
2. **查看备份**：脚本每次运行都会创建带时间戳的备份文件
3. **验证 XBRL 数据**：访问 `https://data.sec.gov/api/xbrl/companyfacts/CIK0001738021.json`
4. **参考 REFACTORING_SUMMARY.md**：包含完整的故障排查指南

---

**Last Updated**: 2026-02-23
**Tested Companies**: CMPX (Compass Therapeutics)
**Tested Years**: 2020-2024
**Status**: ✅ All balance sheet checks passed
