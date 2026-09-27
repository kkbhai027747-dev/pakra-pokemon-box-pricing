# Pokémon box pricing

Standalone project: `pakra-pokemon-box-pricing`. Local workbook utility; no store connection or deployment.

Source: [pakra-cards-system](https://github.com/sreylekcheat-coder/pakra-cards-system), commit `00879b2821c7dce87cb1d44d1d81c0b40e08cea2`, directory `tools/pokemon-box-pricing`, extracted on 2026-09-27. Run the commands below from this repository's root. `SOURCE-MANIFEST.json` retains earlier historical provenance; its old paths are not current file dependencies. This repository split does not deploy the project or replace an existing running task.

中文：本目录按独立项目维护，以下命令从本仓库根目录运行。来源为上述提交及子目录；历史来源清单保留用于追溯。分仓不代表已经部署、连接真实账号或替换原本运行的任务。


## English

**Status:** local-file utility with synthetic tests. This repository copy does not replace the original computer's running task. It reads supplier XLSX quotations, identifies boxed products and gift sets, computes proposed prices, and produces a review workbook and change history. It has no Shopify or WeChat connection.

### Install and test

Requires Python 3.10+ and `openpyxl`. From this folder:

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
.\.venv\Scripts\python -B -m unittest -v
Copy-Item config.example.json config.json
```

The 16 tests construct workbooks in temporary folders. They do not need supplier files, credentials, or a store. Configure the ignored `config.json` before running a collection: set `source.path`, the filename pattern, and the pricing parameters. Relative paths resolve from the config file's folder. No real supplier workbook is included.

```powershell
.\.venv\Scripts\python collect.py
.\.venv\Scripts\python collect.py --file "data/input/example.xlsx"
.\.venv\Scripts\python collect.py --daily
```

`--daily` runs immediately, then at the configured local time while the process stays open; it is not an installed Windows scheduled task. Do not start duplicate daily processes. The optional Windows launchers prefer this project's `.venv`, then `python` on PATH. `run-daily.cmd` delegates to the single launcher.

### Code map and pricing contract

| Area in `collect.py` | Responsibility |
| --- | --- |
| `parse_offer`, `read_offers` | Parse workbook text; flag hidden/ambiguous/duplicate entries |
| `validate_pricing`, `price_offer` | Decimal arithmetic, fee basis, target profit and rounding |
| `quotation_date`, `choose_source` | Select the newest dated quotation; ignore Excel lock files |
| `index_offers`, `changes_since` | Compare products without treating duplicates as reliable quotes |
| `add_sheet`, `save_excel`, `collect` | Build review output and preserve previous output on parsing failure |
| `main` | Load configuration and optionally run the daily loop |

The example retains the current rule: fee = 6% of cost, shipping = CNY 70 per box, target profit = 20% of cost. `cost × 1.26 + 70` is only an approximation: the fee and target profit are each rounded to cents first, then the total selling price is rounded upward to a whole yuan. For example, cost 3.97 gives fee 0.24, target profit 0.79 and selling price 75, not 76. `confirmed: false` marks the result as a trial; it does not block execution. Review these values for your actual source. The alternate `fee_base: "sale"` mode solves for the selling-price-based fee instead.

Output under ignored `output/`: `latest.xlsx`, `latest.json`, and dated history. Chinese result labels are retained to avoid breaking existing review workflows: `已报价` = quoted, `未报价` = no price supplied, `待确认` = needs review, `本版未列出` = absent from this edition. Missing price is not zero or an inventory statement. Old file dates remain visible even when collected today.

### Handoff improvements and limits

- Removed machine-specific launcher paths and duplicate launcher logic.
- Replaced dense duplicate-index expressions and nested change/status conditionals with named helpers and explicit branches. Pricing rules, result keys and existing test expectations are retained.
- Kept pure parsing, pricing and comparison functions directly testable; no framework or new runtime dependency was introduced.
- Formulas are not executed. Image-only quotations, changed layouts, unclear quantities and authenticity are outside this parser's scope. Daily failures retry after a minute; no unattended scheduler is installed here.
- Keep `config.json`, supplier inputs, output reports and task-scheduler state local. See `SOURCE-MANIFEST.json` for provenance. Repository changes do not automatically update the original running project.

## 中文

**状态：**处理本地文件的工具，使用合成数据测试。本仓库副本不会替换原电脑正在运行的任务。脚本读取供应商 Excel，识别盒装与礼盒报价、计算建议售价，并生成供人工核对的表格和变动记录；没有连接 Shopify 或微信。

### 安装与测试

需要 Python 3.10+ 和 `openpyxl`。在本目录按上面的命令创建虚拟环境、安装依赖并运行测试。16 项测试仅在临时目录创建模拟工作簿，不需要真实价表、凭据或店铺。

将 `config.example.json` 复制成 Git 忽略的 `config.json`，填写来源目录、文件名匹配规则和计算参数后再运行。相对路径以配置文件所在目录为基准。单次运行用 `collect.py`，指定文件用 `--file`，每日循环用 `--daily`。

每日模式会先运行一次，之后在电脑本地时间的指定时刻运行，需保持进程开启；它不会安装 Windows 计划任务，也不会自动开机。不要重复启动多个每日进程。Windows 启动器优先使用本项目虚拟环境，再使用 PATH 中的 Python。

### 维护入口与计算规则

`parse_offer/read_offers` 负责解析与待确认标记；`validate_pricing/price_offer` 负责金额与取整；`choose_source` 选择价表；`index_offers/changes_since` 比较变动；`collect` 组织输出；`main` 负责配置和每日循环。修改某一环节时应先读对应函数和已有测试。

示例沿用当前计算参数：手续费为进货价的6%，每盒运费70元，目标利润为进货价的20%。`进货价 × 1.26 + 70` 只是概算；实际先将手续费和目标利润分别四舍五入到分，再将总售价向上取整到元。例如进货价3.97元，手续费0.24元、目标利润0.79元，最终售价75元，而非76元。`confirmed: false` 仅将结果标为试算，不阻止执行；使用前需核对真实业务口径。`fee_base: "sale"` 是另一种按成交价计算手续费的模式。

输出保存在忽略的 `output/`，包含最新Excel/JSON及历史。缺价不等于零价或缺货，文件报价日期不会被采集日期掩盖。输出中文状态和字段保持兼容。

### 本次优化与限制

移除了本机 Python 路径及重复启动逻辑，把重复商品索引和多层条件表达式改为明确函数与分支；保留原计算规则、结果字段和测试预期，没有新增框架。Excel公式不会执行；图片报价、变更后的排版、数量含义及真伪鉴定需另行处理。

真实配置、供应商原表、结果和计划任务留在本地。来源见 `SOURCE-MANIFEST.json`。仓库修改不会自动更新原运行项目，也不会更改店铺价格。
