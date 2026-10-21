# Excel Report Automator

将重复的取数与整理工作封装为函数库：统一读取 Excel / CSV / 网页表格，自动完成字段重命名、类型转换、缺失率统计与异常值标记，单次处理时间由约 40 分钟降至 3 分钟以内。输出标准化 Excel 汇总表与图表，配置驱动，可在新数据上直接复用。

## 功能

- **统一读取入口**：一个函数读取 `.xlsx` / `.xls` / `.csv` / `.html` 文件，也可直接传入网页 URL 读取静态 HTML 表格；批量读取整个文件夹并自动合并。
- **配置驱动的标准化**：字段别名、目标类型、汇总维度、异常规则全部写在 `config.json` 里，换一批同类数据时改配置即可，无需改代码。
- **数据质量检查**：统计每个字段的缺失数量与缺失比例；对金额、数量等数值字段做 IQR（四分位距）离群点标记；支持自定义业务规则（负数、越界等）。
- **类型转换**：自动剥离货币符号（¥ $ €）、千位分隔符和空格后转数值；日期字段智能解析；转换失败的行会记录到报告中。
- **标准化 Excel 输出**：一个工作簿包含四页——`Summary`（汇总指标 + 分组合计 + 柱状图）、`Data`（清洗后明细）、`Quality`（缺失率表）、`Exceptions`（异常明细）。

## 目录结构

```
excel_translate/
├── main.py                  # 命令行入口，串联整个流程
├── data_reader.py            # 统一读取 Excel / CSV / HTML
├── cleaner.py                # 重命名、类型转换、缺失率、异常标记
├── reporter.py               # 多页 Excel 报表与图表
├── generate_sample_data.py  # 生成示例数据（含异常和别名）
├── config.json               # 字段映射、类型、规则配置
├── requirements.txt
├── sample_data/              # 运行 generate_sample_data.py 后生成
└── output/                   # 运行 main.py 后生成 report.xlsx
```

## 快速开始

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 生成示例数据（两个 Excel、一个 CSV、一个 HTML，含异常行和字段别名）
python generate_sample_data.py

# 3. 运行流水线
python main.py --input sample_data --output output/report.xlsx
```

运行结束后打开 `output/report.xlsx`，依次查看 Summary / Data / Quality / Exceptions 四页。

## 配置说明

`config.json` 是适配新数据集的唯一入口：

| 字段 | 作用 |
|---|---|
| `alias_map` | 源字段名 → 标准字段名映射，解决 "销售额 / Revenue / Amount" 同义异名问题 |
| `type_map` | 每个标准字段的目标类型：`numeric` / `date` / `str` |
| `date_col` | 日期列名，用于派生 `Month` 维度 |
| `amount_col` | 需要汇总求和的金额列 |
| `group_cols` | 汇总分组维度，如 `["Category", "Region", "Month"]` |
| `dedup_key_cols` | 去重业务键，留空则按全部非内部列去重 |
| `anomaly_rules` | 业务异常规则，如 `{"name": "neg_amount", "col": "Amount", "condition": "< 0"}` |
| `iqr_cols` | 需要 IQR 离群点检测的数值列 |

换一批同类数据时，只改 `alias_map` 和几个列名就能复用整个流水线。

## 命令行参数

```bash
python main.py --input <文件夹> --output <输出.xlsx> --config <配置.json> \
               --sheet <工作表> --table <序号> --url <网页地址> --partial
```

- `--sheet`：Excel 工作表名或 0 基序号，默认 `0`（第一个）。传 `"0"` 或 `0` 效果相同。
- `--table`：HTML 页面中第几个 `<table>`，默认 `0`（第一个）。
- `--url`：附加读取一个静态网页上的 HTML 表格，与 `--input` 文件夹合并处理。
- `--partial`：有文件读取失败时仍继续生成报表，失败文件列表写入 Summary 页；不加此参数则失败即中止（默认行为，防止用户误以为所有数据已汇总）。

## 输出示例

运行后会生成 `output/report.xlsx`，包含：

- **Summary**：文件数、清洗后行数、金额合计、异常行数、转换失败数；Category / Region / Month 三个分组合计表；一张 Category 柱状图。
- **Data**：清洗后的完整明细，首行冻结。
- **Quality**：每个字段的缺失数量、缺失比例、总行数，按缺失率降序。
- **Exceptions**：触发异常规则或 IQR 离群的行，`_anomaly_flags` 列标明具体命中的规则。

## 设计要点

- **读取后合并，重命名处理重复列**：所有文件读取完毕后集中合并，再统一重命名和类型转换。当不同文件中 "Amount" 与 "Revenue" 这样的同义字段在合并后出现重名列时，`rename_fields` 会将它们合并为一列（第一个非空值优先），并在值冲突的行打上 `_alias_conflict_<col>` 标记。
- **转换失败保留原值，进入 Exceptions**：日期或金额无法解析时，原始值保存到 `_orig_<col>` 列，主列置为 NaN，该行同时在 `_anomaly_flags` 写入 `conv_fail_<col>` 标签，最终出现在 Exceptions 页。
- **缺失率基于真实空值统计**：读取时所有字段保持字符串类型，`rename_fields` 统一将 `"nan"` / `"None"` / `"N/A"` 等哨兵字符串替换为 `None`，再计算缺失率，不会漏报。
- **去重不吞转换失败行**：去重在类型转换之后进行，两条原始值不同但转换后都是 NaN 的行不会被误合并。
- **读取失败默认中止**：有文件无法读取时程序返回错误码 3，防止用户误以为所有数据已汇总；加 `--partial` 后继续运行，失败文件列表写入 Summary 页。
- **业务键去重**：默认按全部数据列去重，建议在配置里指定订单号等业务键，避免误删同日同商品的两笔真实交易。
- **IQR 作为复核提示**：IQR 标记的不是错误，是偏离较远、值得人工看一眼的数据。

## 依赖

- pandas ≥ 2.0
- openpyxl ≥ 3.1
- lxml ≥ 4.9（读取 HTML 表格）
