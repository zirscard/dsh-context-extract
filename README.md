# dsh-context-extract · 上下文萃取与压缩器

把「前置信息萃取与路由 Agent」做成 DSH 插件：agent 直接调用 **`context_extract`** 工具，把超长、多模态的原始文件（财报工作簿、长合同、业务台账）**压成高密度结构化 Markdown + JSON**，过滤冗余、抽取实体、做勾稽校验，供下游 M1–M4 模块用最省 Token 的上下文。

痛点：几十页财报 / 协议原图 / 长文本直接丢给 AI → Token 爆炸、注意力分散、算错数字、漏掉条款。本插件在「读文件」之前先做一层确定性萃取。

```
原始文件 → context_extract → 类型识别 → 关键指标/表格/条款/实体 + 勾稽 alert → 密集 Markdown + JSON
```

## 一、装了之后发生什么

- 注册一个 agent 工具 `context_extract`（所有会话可用）。
- 注入一段系统提示「上下文萃取与压缩器」：当用户丢来超长原始文件时，agent 会先调它，再按「禁止解释 / 按需抽取 / 表格优先 / 条款编号化」四条规则进一步压缩。
- 纯 host 插件：不新增 GUI、不改 DSH 源码、没有 npm 运行时依赖。

## 二、调用方式

```
context_extract(
  path: "/abs/or/relative/财报.xlsx",   # 必填；.xlsx/.xlsm/.csv/.docx/.txt/.md
  targets: ["货币资金","净利润"],        # 可选：只抽指定指标
  outputDir: "..."                       # 可选：报告目录（默认 <输入目录>/.context-extract）
)
```

返回：`type` / `type_label`（财务报表·法律合同·业务台账）、`key_items`（关键指标交叉表）、`tables`（台账表格）、`clauses`（合同条款切分）、`entities`（公司/日期/金额）、`alerts`（勾稽 `<alert>`）、`markdown`（密集压缩文本）、`token_estimate`（≈字符/4）、`report.markdownPath / jsonPath`。

**只读输入，从不修改原文件。**

## 三、能力边界（v0.1）

| 文件 | 处理 |
| --- | --- |
| `.xlsx/.xlsm` | 财务报表 → 关键指标交叉表 + 勾稽；台账 → 表格 dump（前 200 行） |
| `.csv` | 解析为表 |
| `.docx` | 抽取段落 + 表格，合同按条款切分 |
| `.txt/.md` | 文本直读，识别合同/财务标记 |
| `.pdf` | **不支持**（运行时 pdfplumber 损坏、无 pymupdf）；先经 LibreOffice Kit 转文本或 xlsx |

## 四、勾稽校验（自动，仅在财务报表时）

- `balance-sheet`：资产总计 ≠ 负债合计 + 所有者权益合计（容差 0.1%）→ `<alert>`
- `cash-tieout`：BS 货币资金 ≠ CF 期末现金及现金等价物 → `<alert>`

## 五、安装（desktop profile）

```bash
CLI="/Applications/DeepSeek Harness.app/Contents/Resources/runtime/cli/bin/dsh"
"$CLI" plugin --profile desktop add link:<plugin-dir>/dsh-context-extract
```

确认 `~/.dsh/profiles/desktop/package.json` 的 `dsh.profile.bundles` 里有 `dsh-context-extract`；改完在 GUI 工具列表里确认出现 `context_extract`（没出现则重启一次 DSH）。

卸载：`"$CLI" plugin --profile desktop remove dsh-context-extract`（不删报告）。

## 六、运行环境（自动探测）

| 依赖 | 来源 |
| --- | --- |
| Python 3 + openpyxl + python-docx | `~/.dsh/dsh-runtimes/dsh-primary-runtime/dependencies/python` |

可用 `DSH_CONTEXT_EXTRACT_PYTHON` 覆盖 Python 路径。Python 缺失时工具报错并提示用 `load_workspace_dependencies` 检查。

## 七、目录

```
plugins/dsh-context-extract/
├── index.js          # host 插件：注册 context_extract + 系统提示
├── runner.js         # 探测 python，spawn 萃取引擎
├── core/extract.py   # 引擎 CLI（--input/--targets/--out-json/--out-md）
├── test/make_fixture.py   # 生成自测样本
├── cordis.patch.yml
└── package.json
```

## 八、自测

```bash
PY="$HOME/.dsh/dsh-runtimes/dsh-primary-runtime/dependencies/python/bin/python3"
PY test/make_fixture.py                                   # 生成样本财报/台账/合同
PY core/extract.py --input test/fixtures/sample-financial.xlsx \
  --out-json /tmp/extract.json --out-md /tmp/extract.md
```

`sample-financial.xlsx` 应输出 13 个关键指标、2 条 cash-tieout 勾稽提示；`sample-contract.txt` 切出 3 个条款；`sample-ledger.csv` 识别为台账。

## 九、已知限制

- 关键指标依赖中文行标签（资产总计、负债合计、货币资金……）；非中文模板退化为台账/文本。
- 公式单元格若没有缓存值，按空处理（报告里会标注；需要精确值先用 LibreOffice 重算）。
- 条款切分是结构性切分，不生成语义摘要——「条款号 + 核心义务/权利」的一句话说由 agent 按系统提示完成。
- PDF 不支持，需先转换。

## 十、许可

MIT。
