# 你拿到一张表，先不跑公式，先治理它 —— Smart Data Governance VBA

![](./examples/preview.svg)

有一天你打开一张参考库表，往下翻三页就看到问题了：

- **同一个链接出现了 5 次**，每次配的编码还不一样（有些是过时的，有些是打错的）
- **中文名有重复**：`收纳架 / 收纳架-厨房香料架` —— 这俩是同一个产品，后面带"-"的那个更详细
- **英文名大小写乱七八糟**：`Paper Tableware Set / paper tableware set / Party tableware set` —— 应该统一成首字母大写 + LED/CD/USB 这类特殊词保持全大写
- **编码被新记录覆盖了旧的**——17,000 行历史里同一个链接被用过 8 次，最近一次和最早一次给的编码不一样，你不知道该信哪个
- **材质列里混了中英文**：`铝 frame / PVC material` —— "铝 frame" 里 "frame" 是多余英文要删掉，但 "PVC" 是特殊材质缩写要保留
- **链接带一堆无关后缀**：`amazon.example-domain/dp/B0C8B9VQBY-stocklow-recommended` —— 这些后缀是运营打的标签，不影响匹配，匹配前应该清理
- **同款产品的"备用"行没同步编码**：正式行编码改成 9405423100 了，但行号带"备用"后缀的同款柜号还停在 8539520000

`smart-data-governance-vba` 把这些典型的数据治理需求做成了一个 VBA 宏文件里的 **3 个独立模块**，一键跑完，数据表焕然一新。

---

## 🧩 三个模块，各自解决一个具体的数据脏问题

### 模块 1：去重合并（Deduplicate + Merge）

**干的事**：同一链接 + 同一标识符的多行记录 → 合并成一行，同时：
- 中文名按 `/` 分割去重
- 英文名按 `/` 分割去重 + 首字母大写（LED/CD/USB 这类特殊词保持全大写）
- 材质按 `+` 分割去重
- 批次标识符按 `/` 分割去重
- 从后往前删多余行（这样保留的永远是最新那条）

**数字例子**：
> 原始数据 5 行，同一个链接 `amazon.example-domain/dp/B0BY96NYFK` + 同一个编码 `9405423100`，分散在不同行里。
> 
> 模块 1 跑完：合并成 **1 行**，中文名 `天花板射灯/吸顶灯/镜子灯`，英文名 `Wet Room Light/LED Wet Room Light/Mirror Light`（已格式化），材质 `铝/塑料`，批次标识符 `FCIU9781558/OOCU6278411/TIIU6255739`。

### 模块 2：加权评分标准化（Calibrate By Identifier）

**干的事**：参考库里同一个链接在历史上用过多个编码，哪个最可信？

用加权评分算：
```
综合得分 = 频率得分 × 0.6  +  新旧程度得分 × 0.4
```

**数字对比例子**：
> 同一个链接在 LM_Data 历史里用过 3 个编码：
> 
> | 编码 | 出现次数 | 最新行号 | 频率分 | 新旧程度分 | 综合得分 |
> |---|---|---|---|---|---|
> | A | 5 次 | 第 200 行 | 5/8 = 0.625 | 150/200 = 0.75 | 0.625×0.6 + 0.75×0.4 = **0.675** |
> | B | 8 次 | 第 350 行 | 8/8 = 1.0 | 325/350 = 0.93 | 1.0×0.6 + 0.93×0.4 = **0.972** 🏆 |
> | C | 3 次 | 第 400 行 | 3/8 = 0.375 | 400/400 = 1.0 | 0.375×0.6 + 1.0×0.4 = **0.625** |
> 
> → **B 胜出**，写进 Reference 表 D 列。
> 
> 标准化完还会**自动更新同款备用行**：柜号带"备用"后缀、且品名包含正式品名的那些行，同步改成 B 的编码。

### 模块 3：智能字段清洗（Clean Text Fields）

**干的事**：对中文品名、英文品名、材质三个字段分别做定制清洗。

| 字段 | 规则 | 例子 |
|---|---|---|
| **中文品名** | 按 `/` 分割 → **优先保留带"-"的详细版本**（"收纳架-厨房香料架"比"收纳架"好）→ 去重 | 输入：`收纳架/收纳架-厨房香料架` → 输出：`收纳架-厨房香料架` |
| **英文品名** | 按 `/` 分割 → **保留大写字母最多的那个版本**（"Paper Tableware Set" 比 "paper tableware set" 好）→ 格式化 | 输入：`Paper Tableware Set/paper tableware set` → 输出：`Paper Tableware Set` |
| **材质** | 按 `/` 分割 → 每个部分：中文 + 特殊英文缩写（PVC/PP/ABS/LED）保留，普通英文（frame/material）删掉 → "加"和空格统一成"+" → 去掉孤立前后的"+" | 输入：`铝 frame + PVC material + 金属加塑料` → 输出：`铝/PVC/金属+塑料` |

---

## 🔥 适用场景（至少 6 种不同行业）

### 场景 1 — 跨境电商产品参考库治理
历史积累的参考库表（几万行）里链接重复、编码过时、品名混乱 → 一键治理成干净的主数据表。

### 场景 2 — 制造业 BOM（物料清单）标准化
同一个物料号在历史记录里被用过多个供应商编码 → 加权评分选出最可信的编码，自动同步到所有 BOM 行。

### 场景 3 — HR 员工档案去重
员工档案被多个系统维护出重复行（不同系统给不同 ID）→ 去重 + 合并部门/薪资等分散信息。

### 场景 4 — 财务交易记录清洗
多来源交易流水 → 去重 + 合并相似订单 + 标准化交易类型编码。

### 场景 5 — 医疗设备台账治理
历史设备台账里同一个设备出现多次（不同入库时间）→ 合并成一行，保留最新维护记录。

### 场景 6 — 产品规格表规范化
产品规格表被多团队维护出混乱：同一型号多个名字、规格值格式不统一 → 智能去重 + 格式统一。

---

## 🚀 快速上手（4 步）

### 1. 导入代码
VBA 编辑器（Alt + F11）→ 导入 `src/SmartDataGovernance.bas`。

### 2. 准备你的两张表（列布局可调，见 docs/Configuration.md）

| 工作表 | 必需列 | 干什么 | 被哪个模块用 |
|---|---|---|---|
| **LM_Data**（历史记录） | A-F （默认，含链接和标识符） | 历史参考数据 | 模块 1 + 模块 2 |
| **Reference**（标准化目标） | A-J （默认） | 你要治理的那张表 | 模块 1 + 模块 2 + 模块 3 |

### 3. 运行

```
一键跑三个模块:  RunAllGovernance()
只跑去重合并:    RunModule1()
只跑加权标准化:    RunModule2()
只跑字段清洗:    RunModule3()
```

### 4. 看结果
弹出提示框告诉你每个模块处理了多少行。去重会直接删行，标准化会覆盖编码列，清洗会修改文本列——**建议运行前先备份**。

---

## 🎯 三个关键算法一览

### 去重合并 — 三级决策（链接 + 标识符 = 唯一 key）
完全相同的 key → 合并内容（逐列 `/` 或 `+` 分隔去重）→ 从后往前删多余行。

### 加权标准化 — 频率 0.6 + 新旧程度 0.4
从 LM_Data 建分数字典 → 归一化 → 加权选最高分 → 写 Reference 表 → 自动同步同款备用行。

### 智能清洗 — 三个不同策略
中文优先详细版 → 英文优先大写多版 → 材质中英文混合处理（保留 PVC 等特殊缩写）。

---

## 🛠️ 可调常量（源码顶部，不要在文档抄数值）

全部在 `src/SmartDataGovernance.bas` 文件顶部。改完保存即生效。

| 常量 | 干什么 | 默认值 | 什么时候改 |
|---|---|---|---|
| `FREQ_WEIGHT` / `RECENCY_WEIGHT` | 加权评分权重 | 0.6 / 0.4 | 想更信频率还是更信新旧 |
| `linkSuffixes` | 链接后缀清理列表 | `"-stocklow", "-outofstock", "-default"...` | 你的标签后缀不同时 |
| `specialWords` | 英文格式化时保持全大写的词 | `LED, CD, USB, HDMI, CPU, UK, US...` | 加你行业的特殊词 |
| `specialMaterials` | 材质清洗时保留的英文缩写 | `PVC, PP, PE, ABS, PMMA...` | 加你行业的特殊材质 |
| SHEET / COL_* 常量 | 表名和列字母 | Source sheet 是 LM_Data / Reference | 你的表名或列布局不同时 |

完整说明见 `docs/Configuration.md`。

---

## ⚠️ 运行前必知

1. **去重操作不可逆**——运行前建议备份工作簿（Ctrl+S 另存一份）
2. **链接后缀清理规则**是匹配到的第一个后缀就移除——如果你的链接有多层后缀，只去掉第一个
3. **同款产品识别**依赖"柜号备用后缀 + 包含式品名匹配"——你需要在源表里维护好柜号的正式/备用后缀关系

---

## ⚖️ 合规与免责

- 本项目由作者独立 **clean-room 重写**。不引用或拷贝任何前雇主的专有代码或业务文档。
- MIT 许可证仅覆盖作者本人贡献的实现（src/ 目录）。不构成对任何第三方权利的处分。
- 示例数据为完全虚构、中性的演示数据，与任何真实业务无关联。
- 与任何电子表格软件厂商无关联。各类产品名称仅作为常见运行环境提及，不代表背书。

---

## 📦 文件清单

```
smart-data-governance-vba/
├── src/SmartDataGovernance.bas    ← 唯一代码文件，3 个模块 + 18 个函数
├── examples/sample-data.xlsx         ← 虚构演示工作簿
├── docs/Configuration.md             ← 常量说明 + 表结构 + 第二领域示例
├── README.md
├── LICENSE                           ← MIT
├── .gitignore
└── .gitattributes
```

---

## 📅 版本

- **v1.0.0** — 2026-11-15 — 初始发布：去重合并 + 加权评分标准化 + 智能字段清洗

---

## Quick Summary (English)

**smart-data-governance-vba** is a single-file Excel VBA toolkit with
three independent data-governance modules:

- **Deduplicate + Merge** — Rows sharing the same entity link + same
  identifier are merged. CN/EN names, materials, and batch IDs are
  dedup-aware on "/" and "+" separators. Redundant rows are deleted
  bottom-up so the latest record is always kept.
- **Calibrate By Identifier** — Weighted frequency+recency scoring
  (0.6×freq + 0.4×recency) picks the best identifier code for each
  link. Cleans common suffixes from links before matching. Auto-updates
  sibling rows that share a batch-ID suffix pattern + name inclusion rule.
- **Clean Text Fields** — CN names prefer the more detailed (hyphenated)
  variant; EN names keep the variant with the most capital letters,
  then format-first-letter-cap with special words all-capped; material
  strings handle mixed CN/EN (keep PVC/PP/etc., drop generic "frame"),
  normalize separators, strip lone "+" signs.

Designed as a generic data-governance framework — any spreadsheet with
duplicate rows, ambiguous codes, or messy text fields can benefit.

## Maintenance Note

This README is intentionally **Chinese-primary**. The English section is
a concise summary only. All parameters and constants live in source code —
**do not copy numbers into documentation**. Source is the single source
of truth.
