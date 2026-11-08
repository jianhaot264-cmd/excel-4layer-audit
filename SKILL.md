---
name: jt-excel-4layer-audit
description: JT 自建技能。Excel/WPS 表格数据统计分析的端到端全链路交付方法论（中文名：JT Excel 数据统计分析·四层审计），覆盖"梳理数据与任务逻辑 → 口径定义 → 统计计算 → 四层独立验证 → 复盘与剩余风险披露 → 命名与收尾交付"。核心是四层验证阶梯（L1 独立重实现 / L2 Excel+WPS 多引擎强制重算与交叉比对 / L3 结构审计 / L4 不变量与边界），解决"读缓存值 ≠ 真实计算值"这一根本盲区；同一套阶梯也适用于 docx（Word+WPS 文字各自 Fields.Update 后比对域与表格公式结果）。当任务以 .xlsx/.xls/.csv 表格或含表格公式的 .docx 文档为核心——分组汇总与排名、口径统计（求和/计数/去重/日均）、同比环比与占比、数据清洗与质量体检、条件筛选与 Top-K、透视与图表、公式建模、做题作答、报表生成——或用户要求"统计/汇总/分析/算一下/做对/别算错/交付前审一遍/确保数据与事实正确/准确率 100%"时使用。附可执行脚本：run_all.py、selftest.py、audit_xlsx.py、recalc_verify.py、invariant_check.py、make_traps.py、patch_cells.py、fix_formula_tags.py、verify_docx.py、make_docx_field_sample.py。
metadata:
  agent_created: true
  display_name: "JT Excel 数据统计分析·四层审计"
---

# Excel 数据统计分析 · 四层审计

> 覆盖**表格数据统计分析**类任务：从梳理数据与任务逻辑，到统计计算、四层独立验证、复盘披露、完美收尾。
> 适用文件形态：`.xlsx/.xls/.csv`，以及含表格公式的 `.docx`（Word/WPS 文字）。

## 何时**不要**用这个技能

| 场景 | 为什么不适用 |
|---|---|
| 只是把几个数抄进单元格的机械操作 | 直接改更快，跑四层是浪费 |
| 纯展示型需求（做个图、调个色、排序） | 本技能解决的是"算得对不对"，不解决"好不好看" |
| 目标不明、需求还在飘 | 阶段 1 的"口径三问"会先卡住你——**这是设计如此**，先澄清再算 |
| 源数据本身是脏的、且你没权改 | 本技能只能按给定数据算并声明问题，不能替数据源背锅 |
| 需要 Excel 之外的能力（建模、爬数、BI） | 超出本技能边界，路由到对应工具 |

## 概述

把表格任务当成一次**可被复核的统计交付**，而不是一次性的填数。六个阶段，每阶段有明确产出物与放行条件：

| 阶段 | 产出物 | 放行条件 |
|---|---|---|
| 0 路由 | 选定的操作通道 | 明确走哪条工具路径 |
| 1 梳理 | 口径卡（写进表里） | 口径无歧义，或歧义已显式记录 |
| 2 计算 | 计算过程表 + 公式 | 每个数字都能溯源到公式 |
| 3 验证 | 四层验证报告（L1–L4） | 四层全绿，见"什么才算已验证" |
| 4 复盘 | 剩余风险清单 | 每条 WARN 都有解释或已修 |
| 5 收尾 | 命名后的交付文件 + 交付说明 | 原件未动、已呈现、证据可复现 |

**铁律**（任一不满足即不得宣称"已完成"/"已验证"）：

1. **口径先定，且落进表里。** 歧义项要把主口径与备选口径**都**写出来，别只算一种。
2. **不准读缓存值当结论。** `.xlsx` 里的 `<v>` 是上一次保存方写入的缓存；必须至少用**另一个引擎强制重算**一遍（L2）。
   > **实测警告**：L1（`openpyxl(data_only=True)`）和 L3 读的都是缓存，**它们共享同一个盲区**。
   > 受控实验中注入三类陷阱（缓存篡改 / 公式改错 / 越界引用），**L1 全部漏报、L3 只抓到 1 个、L4 全部漏报，只有 L2 三组全中**。
   > **没有 L2 的"验证"不成立。**
3. **原始数据源一个字都不改。** 计算、作答、说明一律新增；改动只允许发生在新增单元格上。
4. **不用新版专属函数。** 接收方可能用旧版 Excel/WPS，`UNIQUE`/`FILTER`/`SORT`/`XLOOKUP`/`SEQUENCE`/`TEXTJOIN`/`IFS`/`MAXIFS` 会 `#NAME?`，并连锁毁掉引用它的结论区。
5. **任何一次修改后必须重跑全部四层验证。** 改一个单元格 = 之前的结论全部失效。这是最容易破功的一条。
6. **表述纪律**：四层全绿才说"已验证"，并同时给出剩余风险。不说"绝对没错"——**本技能不承诺"100% 准确"，只承诺"算得对可证明 + 歧义已穷尽"**。

---

## 阶段 0：路由——选对操作通道

**绝对不要用通用文件工具或自己写脚本去"模仿"表格操作**：

| 任务性质 | 通道 |
|---|---|
| 读取/编辑**已有**表格的内容（分析、汇总、透视、公式、清洗、图表） | 先加载 `tencent-docs-routing`，按其指引委派 `tencent-docs-sheetagent` |
| **确定性小改动**（写已知单元格、设列宽、排序一列、读一个区域） | 加载 `tencent-local-office-edit`，用 `edsdk.py` 调 `sheet_*` |
| **往已有表里补说明行/补单元格**（不需要重算） | `scripts/patch_cells.py`——直接改 XML，**不经过引擎**，因此不会引入公式标签问题，也不抢文件锁 |
| 从零**新建**表格（无源表格文件） | `tencent-docs-sheet-generation` |
| **只做包层修复**（改公式标签、改共享字符串） | `scripts/fix_formula_tags.py` / 直接 `zipfile` — 但**先释放文件锁**，见阶段 5.1 |

要点：

- `file_id` 只能来自 `get_pool_status` / `create_*` / `open_file` 的返回，**不能拿路径当 file_id 猜**。
- 调 `sheet_*` 写值前先用 `sheet_get_used_range` 探明真实边界，不要凭预览或上次读的行数推断。
- 委派 sheetagent 时把**整个**任务一次交出去，不要先自己读一半再委派。

---

## 阶段 1：梳理数据与任务逻辑

> 统计分析的绝大多数错误**不出现在算错，而出现在口径与数据理解**。这一阶段是收益最高的一步。

### 1.0 统计口径三问（先答，再动手）

| 问题 | 为什么致命 |
|---|---|
| **统计谁**（分组键是谁） | 分类列混入异质取值时会静默漏行（如门店代码 `711` 混在文本零售商里） |
| **统计什么**（度量与聚合方式） | 求和 / 计数 / 去重计数 / 均值 / 日均——同一份数据结论完全不同 |
| **按什么范围**（分母与时间窗） | 天数用"自然跨度"还是"有业务的天数"，直接改变日均与排名 |

### 1.1 结构盘点

- 列出全部工作表及角色；记录数据条数、字段、单位、时间跨度；检查隐藏列/隐藏表。
- 用 `sheet_get_used_range` 拿真实边界——**`max_row` 常因格式残留而虚高**（实测数据到 78 行、`max_row` 却是 84），公式引用范围必须按**真实**数据行数写。
- 确认日期系统（1900 / 1904），避免跨系统比较。

### 1.2 数据质量体检

`scripts/invariant_check.py --auto-edge --source <数据源表>` 会自动扫出：

| 检查项 | 处理原则 |
|---|---|
| 完全重复行 | **默认按原表全量统计**，把"去重后的结果"作为备注写出来 |
| 分类列混入异质取值 | 判断是否影响结论；影响就必须与用户确认口径 |
| 空白 / 负值 / 全零列 | 确认是否符合业务含义 |
| 单成员分组 | 这类组的"日均"分母极不稳，口径必须写明 |
| 日期断档 | 分母用"跨度天数"还是"有数据的天数"，必须写明 |

### 1.3 口径卡（产出物）

写成一个文本块放进工作表：

```
口径：金额 = 数量 × 单价；统计天数 = N 天（日期范围 A ~ B）。
备选口径：若按"实际有业务的天数"计算，结果为 …（列出关键差异，注明是否改变名次）。
数据质量说明：第 X/Y/Z 行与第 A/B/C 行完全重复，本次按全量统计；去重后结果变为 …
             「零售商」列有 N 行为门店代码而非文本，若按该列分组需先统一口径。
```

**歧义必须写出来，不要默默选一种。** 排名类问题尤其危险：换口径可能改变名次**顺序**。

---

## 阶段 2：计算

> 统计计算的两个硬要求：**每个数字可溯源到公式**，**每个结论随数据变化而活**。
> 手打一个数字进结论区，就等于埋下一颗"数据改了结论不改"的雷。

### 2.1 工作表布局

- 新建 `作答区`（或 `计算过程`）表，**不动原表结构**。
- 顺序：口径说明 → 明细计算 → 汇总结论 → 图表数据源 → 图表。
- 图表数据源单独成块、整齐降序，图表只引用它。
- 结论区用**跨表引用**指回计算区（如 `=作答区!B20`），保证结论是"活的"。

### 2.2 函数选型

- **白名单（任何版本都安全）**：`SUM` `SUMIF` `SUMIFS` `SUMPRODUCT` `COUNT` `COUNTIF` `COUNTIFS` `ROUND` `RANK` `INDEX` `MATCH` `LARGE` `SMALL` `IF` `IFERROR` `VLOOKUP` `LEFT/RIGHT/MID` `TEXT` `DATE` 系列。
- **黑名单（365/新版 WPS 专属）**：`UNIQUE` `FILTER` `SORT` `SORTBY` `SEQUENCE` `XLOOKUP` `LET` `TEXTJOIN` `IFS` `MAXIFS` `MINIFS` `RANDARRAY` `TAKE` `DROP`。
- 需要"去重清单"时**写静态文字列表**——清单没有计算价值，稳定性远重于自动扩展。
- 经典替代：
  ```
  单条件去重计数  =SUMPRODUCT(1/COUNTIF(范围,范围))
  分组去重计数    =SUMPRODUCT((组=条件)/COUNTIFS(组,组,键,键))
  第 k 名提取     =INDEX(名称区,MATCH(LARGE(值区,k),值区,0)) + =LARGE(值区,k)
  ```
- 硬约束：结论区**不允许出现手打数字**（除排名序号这类标签），否则数据一变就不跟着变。

### 2.3 格式规范

- 金额类 `#,##0.00`；件数/天数/排名 `0`；表头加粗 + 浅底（如 `#DDEBF7`）。
- 说明性长文本放在数据区之外的列。

---

## 阶段 3：四层独立验证 ★ 本技能的核心

> **认知前提：读到的"值"不是"真值"。** `.xlsx` 里 `<v>` 是上一次保存方写入的**缓存**。任何"读缓存 → 比对期望值"的检查，都只证明了写入程序和期望值一致，**没证明 Excel 会算出这个数**。
> 完整方法论见 `references/verification-layers.md`。

### L1 · 独立重实现（differential testing）

用自己的代码从**数据源**重推全部答案，与交付物逐格比对。规范见 `references/audit-checklist.md` 的模板。

必须遵守：换实现路径 / 从数据源读而非从答案区反读 / **按名称查表**比对（不按下标）/ 容差 0.005 / 明细+汇总+结论三处都比。

### L2 · 第二引擎强制重算（N-version 交叉验证）★★ 唯一不可省略的一层

> **为什么它是唯一不可省略的**：L1（`data_only=True`）和 L3 读的都是**缓存**，共享同一盲区。
> 受控实验（注入缓存篡改 / 公式改错 / 越界引用三类陷阱）：**L1 漏报 3/3、L3 漏报 2/3、L4 漏报 3/3，L2 命中 3/3**。
> L2 是整套阶梯中唯一能回答"**引擎重算真的会得到这个数吗**"的一层。

```bash
# 自动探测本机可用引擎
python scripts/recalc_verify.py 成品.xlsx --list-engines

# 双引擎全开（推荐）：Excel + WPS 各自重算，并互相比对
python scripts/recalc_verify.py 成品.xlsx --engine excel,wps --expect expected.json --json r2.json
```

四项比对：① **缓存值 vs 各引擎重算值**（找陈旧/错误缓存）② **引擎之间逐格交叉比对**（暴露引擎差异）③ 引擎重算值 vs 期望值 ④ 重算后错误值扫描。

引擎注册表（脚本内置）：

| key | ProgID | 说明 |
|---|---|---|
| `excel` | `Excel.Application` | Microsoft Office |
| `wps` | `KET.Application` | WPS 表格（ET） |

**两条实战教训**（见 `references/pitfalls.md` P13）：
- WPS 通过 COM 自报名的 `Name` 是 **`"Microsoft Excel"`** —— 判定引擎**不能只看 `Name`**，必须靠 ProgID + 安装路径。否则会出现"以为在验 Excel、其实跑的是 WPS"。
- WPS 可能装在 **非 C 盘**（实测 `E:\WPS Office\...`）。**不要因为 `C:\Program Files\WPS Office` 不存在就断定没装**——用 ProgID 探测才可靠。

**引擎全不可用时的替代**：`pip install formulas`（纯 Python 公式求值库，不需装 Office）/ LibreOffice headless。**若接收方用 WPS，务必用 WPS 也验一遍**——这是消除"接收方引擎差异"这一残留风险的唯一办法。

### L3 · 结构与静态审计

```bash
python scripts/audit_xlsx.py 成品.xlsx --orig 原始文件.xlsx --expect expected.json --json r3.json
```

检查项（**唯一权威清单以本节为准**，其它地方只引用不再复述，避免数字漂移）：

1. zip 完整性（`testzip`）与每个 xml/rels 可被 `ElementTree` 解析
2. 工作簿主部件存在
3. `<f>` 标签内是否有前导等号
4. 是否残留 `t="array"` 数组标记
5. **公式单元格是否都带 `<v>` 缓存值**（缺了部分阅读器显示空白）
6. 是否使用了 365/新版 WPS 独占函数（按函数名边界匹配，不会把 `COUNTIFS(` 误判成 `IFS(`）
7. 全表 `#REF!` `#VALUE!` `#NAME?` `#DIV/0!` `#N/A` `#NULL!` `#NUM!` 扫描
8. 分类列类型一致性（纯表格里混入数值的分类列会**静默漏行**）
9. 公式结构一致性（归一化坐标后分组，找离群单元格）
10. 真实数据边界 vs 公式引用区间
11. 图表：series 引用区域与缓存点数
12. 数字格式抽查（只盯"含公式的输出列"）
13. 若给了 `--orig`：逐格 diff，确认原表与题干**零改动**、新增单元格可枚举
14. 若给了 `--expect`：按 JSON 清单定点校验期望值

退出码 0/1，可接自动化。

### L4 · 不变量与边界（property-based）

```bash
python scripts/invariant_check.py 成品.xlsx --spec invariants.json --auto-edge --source 数据原表
```

检查"必须成立的性质"而非具体数字：排名完整性 / 降序性 / 配平 / **结论区 Top-K 等于全量池前 K 名**（端到端不变量，最值钱）/ 并列卡边界 / 硬编码数字。`--spec` 的写法见脚本头部注释。

### 什么才算"已验证"

- [ ] L1 全部答案逐格吻合，无例外
- [ ] L2 每个可用引擎：缓存=重算（0 处不一致）；**多引擎之间逐格 0 处差异**；重算=期望（全命中）；0 错误值
- [ ] L3 0 FAIL，每条 WARN 有解释
- [ ] L4 0 FAIL，边界 WARN 已写进交付物
- [ ] 交付说明含口径 / 备选口径 / 数据问题 / 剩余风险
- [ ] 记录文件 SHA256，保证"验证的就是交付的那个文件"，并确认引擎以只读方式打开后哈希未变

**顺序纪律**：修改后重跑四层。改一个单元格就重来。

---

## 阶段 4：复盘与剩余风险

- 逐条处理 WARN：能修的修，不能修的**写进交付说明**。
- 显式列出残留风险（口径理解偏差 / 数据源本身有误 / 接收方引擎差异 / 题面隐含前提）。清单见 `references/verification-layers.md` 第五节。
- **表述纪律**：
  - ✅ "已验证：四层检查全绿；剩余不确定性是 X"
  - ❌ "准确率 100%，绝对没错"、"应该没问题"

---

## 阶段 5：收尾交付

### 5.1 写盘前释放文件锁（最容易卡住的一步）

只要表格引擎里还有该文件的实例（**包括宿主预览自动注册的实例**），磁盘文件就是 deny-write：`os.replace` 报 `WinError 183/32`，`open(path,'r+b')` 报 `PermissionError 13`。

```
1. get_pool_status                      # 找出所有实例
2. close_file file_id=<id> force=true   # 逐个关闭
3. 删除同目录残留的 ~$xxx.xlsx
4. 带重试落盘（间隔 3 秒）
5. 仍失败才 shutdown force=true         # 最后手段
```

> ⚠️ `shutdown` 之后服务**不会自动恢复**，只能等宿主按需重启。**能只 `close_file` 就别 `shutdown`。**

`scripts/patch_cells.py` / `fix_formula_tags.py` 自带重试与明确的失败提示，可直接用。

### 5.2 命名与保存

- **永远另存为新文件，不覆盖原件**；原件只读时先复制到工作区再改。
- 命名遵循接收方约定（如"文件+姓名"）；未明说时用 `原文件名-姓名.xlsx`，并提示可自行改名。
- 保存后不要主动 `close_file`（除非为了释放锁）。

### 5.3 交付说明（必须包含）

1. **结论**：每个指标的最终答案，表格列出
2. **口径**：主口径 + 备选口径及其差异
3. **发现的数据问题**：重复行、类型混杂、异常值，及对结论的影响
4. **原文改动情况**：哪些是新增、哪些一动没动
5. **剩余风险**：无法消除的不确定性
6. **验证证据**：四层检查结论 + 文件 SHA256
7. **文件路径与命名**，并提示能否改名/调整排版

最后用 `present_files` 呈现成品。

---

## 附：文档类（docx）验证

同一套思路平移到 Word/WPS 文字。**盲区是同一个**：文档里的域（Field）结果、表格公式
（`=SUM(ABOVE)` 这类）的结果，也是"上一次保存方写进文件"的缓存，不代表重算后的值。

```bash
python scripts/verify_docx.py 文件.docx --list-engines
python scripts/verify_docx.py 文件.docx --engine word,wps --expect expected.json --json r.json
```

| key | ProgID | 说明 |
|---|---|---|
| `word` | `Word.Application` | Microsoft Word |
| `wps` | `KWPS.Application` | WPS 文字 |

它做四件事：① 两引擎各自 `Fields.Update()` 强制重算（**文档版的"全量重算"**，含表格公式）
② 读出非空段落 / 各表格单元格 / 各域结果 ③ 跨引擎逐项比对 ④ 与期望值清单比对。

引用写法：`para:3`（第 3 个非空段落）、`table:1:R2C3`（第 1 表第 2 行第 3 列）、`field:2`（第 2 个域的结果）。

**验证有效性的判据**：把文件里的域缓存**故意写成错值**，看引擎能否算回正确值。
用 `python scripts/make_docx_field_sample.py` 一键生成该样本（`=SUM(ABOVE)` 缓存写成 `999999`，
正确的两引擎都应算回 `60`；若都报 `999999`，说明它们只读了缓存，验证无效）。

> 注意：`KWPS.Application` 同样把 `Name` 自报成 `"Microsoft Word"`——引擎判定规则与表格侧完全一致，
> 见 `references/pitfalls.md` P13。

---

## 资源

| 文件 | 用途 |
|---|---|
| `scripts/audit_xlsx.py` | L3 结构审计：全维度检查（清单见"阶段 3 · L3"），支持 `--orig` 逐格 diff / `--expect` / `--json` |
| `scripts/run_all.py` | **一键跑完四层**并产出证据包（`report.md` + `report.json` + 各层原始输出），含文件 SHA256 与改动检测。退出码 0=全过 / 1=有 FAIL / **3=有层未执行（部分验证）** / 2=前置不满足 |
| `scripts/selftest.py` | **技能自检**：在临时目录造合成样本，把所有脚本的正常+异常路径跑一遍（45 项）。换机器/换项目后**先跑它**，证明工具链是好的，再验真实交付物 |
| `scripts/make_docx_field_sample.py` | 生成 docx 域更新自检样本（缓存故意写错），用于验证 `verify_docx.py` 的"重算"真的生效 |
| `scripts/recalc_verify.py` | L2 多引擎强制重算（**表格**）：Excel + WPS 各自重算并互相比对，`--list-engines` 探测可用引擎，退出码 2 = 无引擎 |
| `scripts/verify_docx.py` | **文档**版孪生：Word + WPS 文字各自 `Fields.Update()` 重算，段落/表格/域跨引擎比对 |
| `scripts/invariant_check.py` | L4 不变量与边界：7 类 spec 检查 + `--auto-edge` 自动体检 |
| `scripts/patch_cells.py` | 直接改 XML 写单元格（不经过引擎，不抢锁，写入后自校验失败自动回滚） |
| `scripts/fix_formula_tags.py` | 归一化公式标签（去前导等号 + 消除数组标记） |
| `scripts/make_traps.py` | **边界实验**：注入三类陷阱实测各层能否捕获（自带 demo 样本，不需外部文件） |
| `scripts/release_check.py` | **发布前全检**：一条命令跑 8 组检查（换行符 / 硬编码路径 / 语法 / frontmatter / 自检 / 边界实验复现 / 打包 / CLI dry-run），输出"能不能发"的结论 |
| `scripts/make_publish_package.py` | 生成 SkillHub 发布包（增补 5 个 SkillHub 必需字段 + 附公开 README + 自动 dry-run） |
| `references/verification-layers.md` | **验证阶梯方法论**：L0–L4、DoD、表述纪律、残留风险、可复现证据 |
| `references/boundary-experiment.md` | **实证：L1/L3/L4 会漏报、只有 L2 全中** —— 为什么 L2 不可省略 |
| `references/pitfalls.md` | P1–P20 已知坑与修法（**动手前先读**） |
| `references/audit-checklist.md` | 逐项判定标准 + 独立重算脚本模板 + 交付说明模板 |
| `references/monetization.md` | 变现设计：三层模型、定价框架、付费交付物、SkillPay 接法、冷启动路径 |
| `references/retrospective.md` | 复盘：撞名/猜规则/装 vs 发 三条走错路的判别式 |
| `references/public-readme.md` | 面向市场详情页的公开 README（发布包会把它作为 README.md） |

### 发布 / 更新本技能（作者用）

```bash
# 一条命令：全检 + 打包 + dry-run，通过才允许发
python scripts/release_check.py --with-experiment

# 正式发布（Token 只在你自己的终端传，不要落到聊天记录里）
MSYS_NO_PATHCONV=1 <带 requests 的 python> \
    <home>/.skillhub/skills_store_cli.py publish <包>.zip \
    --token skh_xxx --changelog "本次改了什么"
```

> ⚠️ **CLI 只能发布「技能包」，不能发布「收费」**。
> 定价、Pay Skill 计费、支付宝 serviceId 绑定**只能在 skillhub.cn 网页端做**。
> 同 slug 再发一次 = 版本更新（不新建 listing），记得升 `version:` 并写清 changelog。

### 一条命令跑完

```bash
python scripts/run_all.py 成品.xlsx \
       --orig 原始.xlsx --expect expected.json --invariants invariants.json \
       --source 数据源表 --l1 任务专属重算.py --engines excel,wps --out ./验证报告
```

### 新机器 / 新项目第一步

```bash
python scripts/selftest.py                # 45 项工具链自检，不触碰任何真实文件
python scripts/selftest.py --with-engines # 额外验证 Excel/WPS 引擎通道
```

**为什么必须有自检**：检查脚本自己也可能悄悄失效——本次实测就发现
`fix_formula_tags.py` 在最常见的 `<f>=…` 形式上**一个都修不到**（正则要求 `<f` 后有空格），
而它报的是"0 处需修正，文件已合规"，看起来一切正常。没有自检，这种错误会一直带着走。
