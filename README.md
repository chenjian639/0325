# 文献计量数据词频分析处理

## 项目概述

对 2012–2018 年 Web of Science 收录的知识组织相关论文进行词频分析，按 **WoS 5 大研究领域 → 国别 → 年份** 分组统计关键词频率。

领域划分依据：[Web of Science Research Areas 官方分类](https://webofscience.zendesk.com/hc/en-us/articles/38543541713169-Research-Areas)（154 个 Research Area，归入 5 个大类）。

## 数据来源

`0325/` 目录下共 7 个 Excel 文件（2012–2018）：

| 文件 | 年份 | 大小 |
|------|------|------|
| 四张sheet_国家_颜色2012.xlsx | 2012 | 2.3 MB |
| 四张sheet_国家_颜色2013.xlsx | 2013 | 3.0 MB |
| 四张sheet_国家_颜色2014.xlsx | 2014 | 3.0 MB |
| 四张sheet_国家_颜色2015.xlsx | 2015 | 3.5 MB |
| 四张sheet_国家_颜色2016.xlsx | 2016 | 3.6 MB |
| 四张sheet_国家_颜色2017.xlsx | 2017 | 3.7 MB |
| 四张sheet_国家_颜色2018.xlsx | 2018 | 3.7 MB |

每个 Excel 文件包含 4 个 Sheet：

| Sheet 名 | 含义 | 2012 年记录数 |
|----------|------|--------------|
| ontology | 本体研究 | 1,813 |
| KG | 知识图谱 (Knowledge Graph) | 37 |
| LinkedData | 关联数据 (Linked Data) | 393 |
| Thesaurus | 叙词表研究 | 55 |

### 关键字段

| 字段名 | 说明 | 示例 |
|--------|------|------|
| Year | 年份 | 2012 |
| country | 国家全称 | Peoples R China |
| country_code | 国家代码 | CN |
| Keywords | 原始关键词（无分隔符拼接） | `Author KeywordsOntologyTwo-layer ontology model` |
| Categories/ Classification | 研究领域分类 | `Research AreasComputer Science Citation Topics...` |

### 数据量

- 有效记录：**19,197** 条
- 缺失记录：**3,171** 条（缺国家 718 条，缺关键词 2,062 条，污染关键词 391 条）

---

## 核心挑战与处理方案

### 1. 关键词无分隔符拼接

原始 `Keywords` 字段示例：

```
Author KeywordsOntologyTwo-layer ontology modelOntology building method
```

关键词之间无任何分隔符，仅通过 **小写→大写边界**（驼峰命名法）作为切分点。

**处理流程：**

```
原始字符串
  ↓ Step 1: 分离 Author Keywords 与 Keywords Plus
  ↓ Step 2: 驼峰边界切分 (lowercase→uppercase)
  ↓ Step 3: 词频词典 + NLTK 英语词表 辅助再切分（9,914 个长Token被拆分）
  ↓ Step 4: 去重合并，标准化清洗（去掉不成对括号、chevron_right污染等）
  ↓ 最终关键词列表
```

### 2. 污染数据检测

部分记录的 Keywords 字段被错误写入了 Categories/Classification 数据（以 `Research Areas` 开头或含 `Citation Topics` / `chevron_right`）。共检测出 **391 条**，归入缺失数据处理，不参与词频统计。

### 3. Keywords Plus 处理

约 43% 的记录包含 `Keywords Plus` 部分。当前版本 **仅统计 Author Keywords**，Keywords Plus 保留在原始数据映射表中供参考。

### 4. 研究领域提取与归并

`Categories/ Classification` 字段格式：

```
Research AreasComputer Science Citation Topics4 Electrical Engineering...
Research AreasComputer ScienceLinguistics Citation Topics...
```

使用 WoS 官方 **154 个 Research Area** 进行贪婪最长匹配，然后按 Zendesk 官方分类归入 **5 个大类**：

| 大类 | 包含子领域数 | 本文数据论文人次 |
|------|-------------|-----------------|
| Arts & Humanities | 15 | 384 |
| Life Sciences & Biomedicine | 76 | 3,333 |
| Physical Sciences | 17 | 1,135 |
| Social Sciences | 25 | 1,637 |
| Technology | 21 | 15,255 |
| Unknown | — | 1 |

一篇论文可能标记多个 Research Area，其关键词会贡献到 **所有** 匹配大类的输出中。

### 4. 列位置不一致与表头损坏

不同年份、不同 Sheet 的列位置有差异。2018 年 KG Sheet 表头损坏，数据出现在表头行。

**处理方案：**
- 优先按列名匹配（header-name-based lookup）
- 未匹配到时使用已知位置回退（positional fallback）
- 2018 KG 自动回退并记录警告

---

## 脚本说明

**主脚本：** `process_bibliometrics.py`

### 运行方式

```bash
# 默认处理 2012-2018
python process_bibliometrics.py

# 指定年份和目录
python process_bibliometrics.py --years 2012 2013 2014 --output-dir ./my_output

# 自定义输入目录
python process_bibliometrics.py --input-dir ./data --output-dir ./output
```

### 依赖

- `pandas` — Excel 读写
- `openpyxl` — Excel 文件解析
- `nltk` — 英语单词表（可选，用于关键词辅助切分，未安装时会降级处理）

### 处理流程

```
Phase 1 (Pass 1): 读取所有文件，驼峰切分 → 构建词频词典
    ↓
Phase 2: 词典构建 (词频 ≥ 2, 长度 ≥ 3)
    ↓
Phase 3 (Pass 2): 使用词典 + NLTK 英文词表 → 再切分长Token
    ↓
Phase 4: 按 (大类, 国别, 年份) 聚合 → 词频统计
    ↓
Phase 5: 输出 txt + Excel
```

---

## 输出文件

### 目录结构

```
output/
├── by_field/                             ← 按 5 大类的 txt 文件
│   ├── Technology.txt                    ← 1.5 MB（最大）
│   ├── Life Sciences & Biomedicine.txt   ← 528 KB
│   ├── Social Sciences.txt               ← 184 KB
│   ├── Physical Sciences.txt             ← 163 KB
│   ├── Arts & Humanities.txt             ← 32 KB
│   └── Unknown.txt                       ← 76 KB
├── 原始数据映射表.xlsx                   ← 19,588 条记录的关键词映射
├── 缺失数据记录.xlsx                     ← 2,780 条缺失字段的数据记录
├── 领域列表.txt                          ← 当前输出的 6 个大类
├── WoS_Research_Areas_完整列表.txt        ← WoS 官方 154 个 Research Area（5大分类）
└── WoS完整学科领域列表.txt               ← WoS 官方 254 个 Subject Categories（参考）
```

### txt 文件格式

```
【Technology】
(Peoples R China)
(2012)
Ontology: 45
Semantic Web: 32
knowledge organization: 11
...
(2013)
Ontology: 52
...
(US)
(2012)
Ontology: 38
...
【Life Sciences & Biomedicine】
...
```

- 第一层：`【大类名】`
- 第二层：`(国别)`
- 第三层：`(年份)`
- 第四层：`关键词: 频次`，按频次降序排列

### 原始数据映射表（Excel）

| 列名 | 说明 |
|------|------|
| record_id | 记录编号 |
| source_year | 来源年份 |
| source_sheet | 来源 Sheet |
| source_row | Excel 行号 |
| country | 国家 |
| country_code | 国家代码 |
| research_areas | 匹配的 WoS 子领域（\| 分隔） |
| broad_categories | 归属的 WoS 大类（\| 分隔） |
| author_keywords | 解析的作者关键词（\| 分隔） |
| kwplus_keywords | 解析的 Keywords Plus（\| 分隔） |
| all_keywords_merged | 合并去重后的所有关键词（\| 分隔） |
| raw_keywords | 原始 Keywords 字段 |
| raw_categories | 原始 Categories/Classification 字段 |

> 查错方式：通过 `source_year` + `source_sheet` + `source_row` 可定位到原始 Excel 的具体行。

### 缺失数据记录（Excel）

| 列名 | 说明 |
|------|------|
| source_year | 来源年份 |
| source_sheet | 来源 Sheet |
| source_row | Excel 行号 |
| reason | 缺失原因（missing_country / missing_keywords / contaminated_keywords） |
| raw_keywords | 原始关键词（如有） |
| raw_categories | 原始分类（如有） |
| raw_country | 原始国家（如有） |

---

## 处理统计数据

### 总体

| 指标 | 数值 |
|------|------|
| 输入文件数 | 7 |
| 有效记录 | 19,197 |
| 缺失记录（已排除） | 3,171 |
| 唯一关键词 | 37,722 |
| 词频条目总数 | 87,639 |
| 输出文件数 | 6（5 大类 + Unknown） |

### 按年分布

| 年份 | 有效记录 | 缺失记录 |
|------|---------|---------|
| 2012 | 2,016 | 331 |
| 2013 | 2,548 | 425 |
| 2014 | 2,513 | 451 |
| 2015 | 3,050 | 417 |
| 2016 | 3,103 | 435 |
| 2017 | 3,146 | 416 |
| 2018 | 2,821 | 696 |
| **合计** | **19,197** | **3,171** |

### 按 Sheet 分布

| Sheet | 记录数 |
|-------|--------|
| ontology | 13,925 |
| LinkedData | 4,103 |
| KG | 608 |
| Thesaurus | 561 |
| **合计** | **19,197** |

### 5 大类分布（论文人次）

| 大类 | 论文人次 |
|------|---------|
| Technology | 15,255 |
| Life Sciences & Biomedicine | 3,333 |
| Social Sciences | 1,637 |
| Physical Sciences | 1,135 |
| Arts & Humanities | 384 |
| Unknown | 1 |

---

## 已知局限

1. **小写关键词无边界切分**：如 `networkontology` 无大小写变化，无法通过驼峰规则切分。已通过词频词典和 NLTK 英语词表辅助切分，部分低频组合可能仍保留为长Token。

2. **研究领域匹配**：仅能匹配 WoS 官方 154 个 Research Area 名称。1 条记录的 Categories/Classification 字段不含有效领域信息，归入 `Unknown`。另有 391 条记录的 Keywords 字段被 Categories 数据污染，归入缺失数据。

3. **国家名称非标准**：Scotland、Wales、North Ireland 使用名称本身作为代码（非 ISO 标准），已保留原样。

---

## 领域分类参考

| 文件 | 说明 |
|------|------|
| `output/WoS_Research_Areas_完整列表.txt` | WoS 官方 154 个 Research Area，按 5 大类分列 |
| `output/WoS完整学科领域列表.txt` | WoS 官方 254 个 Subject Categories（SCIE/SSCI/AHCI），更细粒度 |
| `output/领域列表.txt` | 本次输出的 6 个大类名称 |

---

## 文件清单

| 文件 | 说明 |
|------|------|
| `process_bibliometrics.py` | 主处理脚本 |
| `output/` | 输出目录 |
| `0325/` | 原始数据目录（2005–2024 全部年份 Excel + country_code 统计） |
| `README.md` | 本文件 |
