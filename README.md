# excel-4layer-audit

Excel / WPS 表格数据统计分析的「四层审计」技能包源码。

> 这是一套 WorkBuddy Skill 的源码（SKILL.md + 参考文档 + 可执行脚本），
> 不是一个可直接 pip install 的独立 Python 库。

## 解决什么问题

`.xlsx` 单元格里的 `<v>` 是上一次保存方写入的**缓存值**，不是公式的当前真实计算结果。
用 `openpyxl` 的 `data_only=True` 读到的数字可能已经过期，直接拿来当结论会出错。

本技能把表格任务当作一次**可被复核的统计交付**，用四层独立验证消除这个盲区：

| 层 | 手段 | 拦住什么 |
|---|---|---|
| L1 | 独立重实现（从源数据重新推导，逐格比对） | 同一实现里的系统性错误 |
| L2 | Excel / WPS 多引擎强制重算 + 交叉比对 | 缓存值与真实计算值不一致 |
| L3 | 结构审计（公式、区域、合并单元格、口径卡） | 结构与口径层面的错配 |
| L4 | 不变量与边界检查（配平、排名连续、极值） | 逻辑上不可能成立的结果 |

同一套阶梯也适用于含表格公式的 `.docx`（Word 与 WPS 文字各自 `Fields.Update` 后比对域结果）。

六阶段流程：路由 → 口径定义 → 计算 → 四层验证 → 复盘与剩余风险披露 → 命名交付。
放行条件与铁律见 [SKILL.md](SKILL.md)。

## 目录结构

```
SKILL.md                    技能定义、触发条件、六阶段流程、放行铁律
references/
  audit-checklist.md        审计检查清单
  boundary-experiment.md    边界实验设计
  monetization.md           商业化路径设计（公开版已泛化）
  pitfalls.md               踩坑记录
  public-readme.md          公开版说明稿
  retrospective.md          复盘与判别式
  verification-layers.md    四层验证的实现细节
scripts/
  run_all.py                全流程编排
  selftest.py               自测
  audit_xlsx.py             L3 结构审计
  recalc_verify.py          L2 多引擎重算比对
  invariant_check.py        L4 不变量校验
  make_traps.py             合成陷阱样本生成
  patch_cells.py            单元格修补
  fix_formula_tags.py       公式标签修复
  verify_docx.py            docx 域与表格公式校验
  make_docx_field_sample.py docx 样本生成
  make_publish_package.py   发布包打包
  release_check.py          发布前检查
```

## 运行

脚本依赖以各文件 `import` 为准；L2 的多引擎重算需要本机安装 Excel 或 WPS，
Windows 下通过 ProgID 探测，且 WPS 可能装在非系统盘，不要按固定路径硬判。

建议先跑 `python scripts/selftest.py` 验证环境与合成样本，再跑 `run_all.py`。

## 公开版说明

本仓库是**公开版**，相对本地原始版本做了最小必要脱敏：

- 移除具体机构名称与在审项目名称，改为通用场景描述；
- 移除含本机用户名的绝对路径示例，改为占位符；
- 技术内容、方法论、脚本逻辑与验证结论均未改动。

边界实验使用脚本自带的合成样本，不含任何真实业务数据。
未脱敏的原始版本仅保存在本地，未随本仓库分发。

## 使用边界

本技能解决「算得对不对」，不解决「好不好看」。源数据本身脏且无权修改时，
技能只能按给定数据计算并显式声明问题，不替代数据源质量责任。
