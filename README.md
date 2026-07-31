# 网络毒理学靶点预测脚本套件

一站式网络毒理学/药理学研究工具集，覆盖 4 大权威数据库靶点查询 + 4 大化合物靶点检索/预测平台 + 分子属性计算与结构处理。

终端交互式操作，零编程基础上手，结果一键导出 Excel。

---

## 项目结构

```
├── 疾病靶点/                              # 4 款 — 疾病名称 → 靶点查询
│   ├── TTD_疾病靶点搜索.py                 # ① TTD 疾病靶点搜索
│   ├── DisGeNET_疾病靶点检索.py            # ② DisGeNET 疾病靶点检索
│   ├── OpenTargets_疾病靶点检索.py          # ③ Open Targets 疾病靶点检索
│   └── PharmGKB_基因靶点搜索.py            # ④ PharmGKB 基因靶点搜索
├── 分子靶点/                              # 4 款 — 化合物结构 → 靶点预测
│   ├── ChEMBL_化合物靶点检索.py            # ⑤ ChEMBL 化合物靶点检索
│   ├── SuperPred_靶点预测.py               # ⑥ SuperPred AI 靶点预测
│   ├── SwissTarget_靶点预测.py             # ⑦ SwissTargetPrediction 靶点预测
│   └── TargetNet_靶点预测.py               # ⑧ TargetNet QSAR 靶点预测
├── 分子处理/                              # 1 款 — 本地分子属性计算
│   └── 小分子结构处理工具.py                # ⑨ 分子属性计算 & 结构处理
├── requirements.txt                       # 共享依赖清单
├── 一键安装依赖.bat                        # 一键安装所有依赖
└── README.md                              # 本文件
```

---

## 快速开始

### 环境要求

- **操作系统**: Windows 10 / 11（64 位）
- **Python**: 3.9 及以上
- **网络**: 需要（除脚本 9 外均需联网访问数据库/预测平台）

### 安装依赖

**方式一（推荐）**：双击 `一键安装依赖.bat`，等候 3-10 分钟自动完成。

**方式二（手动）**：

```bash
pip install -r requirements.txt
playwright install chromium
```

### 运行脚本

```bash
python 疾病靶点/TTD_疾病靶点搜索.py
python 分子靶点/ChEMBL_化合物靶点检索.py
python 分子靶点/SuperPred_靶点预测.py
```

所有脚本均采用终端交互菜单，按提示操作即可。

---

## 脚本详情

### 一、疾病靶点类（4 款） — `疾病靶点/`

输入疾病名称，从权威数据库查询关联靶点。

#### 1. TTD 疾病靶点搜索 — `TTD_疾病靶点搜索.py`

| 项目 | 说明 |
|------|------|
| 数据源 | TTD (Therapeutic Target Database) — 全球最大的治疗靶点数据库 |
| 输入 | 疾病英文名称（如 Alzheimer、Diabetes）或 ICD-11 编码 |
| 输出 | Excel (.xlsx)，含靶点 ID、基因名、相关疾病、代表药物 |
| 亮点 | 基本模式（秒级）与详细模式（含 UniProt/染色体位置/功能描述） |

#### 2. DisGeNET 疾病靶点检索 — `DisGeNET_疾病靶点检索.py`

| 项目 | 说明 |
|------|------|
| 数据源 | DisGeNET — 全球最大的基因-疾病关联知识平台 |
| 输入 | 疾病英文名称 或 疾病 ID（UMLS/MONDO/OMIM/MeSH 等） |
| 输出 | Excel (.xlsx)，含 3 个 Sheet（基因疾病关联/TOP100/汇总统计） |
| 亮点 | 专业评分体系（score/DSI/DPI/pLI/EI）· 可下载全部疾病列表（≤10,000 条）· 独有 Disease Ontology & MONDO 本体下载 |
| 注意 | 需申请 DisGeNET 免费学术 API Key（脚本内置申请引导） |

#### 3. Open Targets 疾病靶点检索 — `OpenTargets_疾病靶点检索.py`

| 项目 | 说明 |
|------|------|
| 数据源 | Open Targets Platform (EMBL-EBI + GSK + Sanger 等联合项目) |
| 输入 | 疾病名称（支持中英文）或疾病 ID（MONDO/EFO） |
| 输出 | Excel (.xlsx)，含 3 个 Sheet，每种证据类型独立成列 |
| 亮点 | 8 维度证据评分（遗传关联/体细胞突变/RNA 表达/通路/已知药物/动物模型/文献/临床）· 多疾病批量查询 · 12 种常见疾病快捷入口 |

#### 4. PharmGKB 基因靶点搜索 — `PharmGKB_基因靶点搜索.py`

| 项目 | 说明 |
|------|------|
| 数据源 | PharmGKB (NIH 资助) — 药物基因组学权威数据库 |
| 输入 | 疾病或药物关键词（如 Diabetes、Aspirin） |
| 输出 | Excel (.xlsx)，基因符号列表 |
| 亮点 | 自动读取总数 · 智能翻页 · 请求重试与速率控制 · 去重排序 |

### 二、分子靶点类（4 款） — `分子靶点/`

输入化合物结构，通过数据库活性数据或 AI 平台预测靶点。

#### 5. ChEMBL 化合物靶点检索 — `ChEMBL_化合物靶点检索.py`

| 项目 | 说明 |
|------|------|
| 数据源 | ChEMBL (EMBL-EBI) — 全球最大开源生物活性数据库 |
| 输入 | SMILES 结构式 或 化合物名称（自动识别） |
| 输出 | 合并 Excel + 每个化合物单独 CSV |
| 亮点 | 3 种搜索模式（精确/相似/子结构）· 8 线程并发获取基因详情 · 可设活性阈值与物种过滤 · Excel/CSV 批量输入 + GUI 弹窗选列 |

#### 6. SuperPred 靶点预测 — `SuperPred_靶点预测.py`

| 项目 | 说明 |
|------|------|
| 平台 | SuperPred (德国 Charité 医学院) — 基于机器学习的靶点预测 |
| 输入 | 化合物 SMILES（支持单化合物/Excel 批量/CSV 批量） |
| 输出 | 每个化合物单独 CSV + 合并 Excel + 汇总报告 + 失败列表 |
| 亮点 | Playwright / Selenium 双引擎 · GUI 弹窗选列 · JS API + DOM 双路径结果提取 · 防封机制 · 5 个内置化合物快速体验 |

#### 7. SwissTargetPrediction 自动化预测 — `SwissTarget_靶点预测.py`

| 项目 | 说明 |
|------|------|
| 平台 | SwissTargetPrediction (瑞士生物信息学研究所 SIB) — 高精度靶点预测 |
| 输入 | Excel (.xlsx) 或 CSV 文件，含 SMILES 列 |
| 输出 | 每个化合物单独 CSV + 合并 Excel + 汇总报告 |
| 亮点 | 7 步交互式配置向导 · RDKit SMILES 合法性预检 · 人/小鼠/大鼠 3 物种 · 强大防封机制 |

#### 8. TargetNet QSAR 靶点预测 — `TargetNet_靶点预测.py`

| 项目 | 说明 |
|------|------|
| 平台 | TargetNet (中南大学) — 基于 QSAR 模型的靶点预测 |
| 输入 | SMILES 字符串 或 .smi 文件 或 Excel 文件 |
| 输出 | 每个化合物单独 CSV + 合并 Excel（csv/xlsx/txt 可选） |
| 亮点 | 7 种分子指纹可选 · 4 种筛选指标 + 可调阈值 · 快速/手动双模式 · 智能自动分批 · GUI 弹窗选列 |

### 三、分子处理类（1 款） — `分子处理/`

#### 9. 小分子结构处理工具 — `小分子结构处理工具.py`

| 项目 | 说明 |
|------|------|
| 依赖 | 纯本地计算（RDKit + OpenBabel），无需网络 |
| 输入 | Excel (.xlsx) / SDF / MOL / MOL2 / PDB，支持文件夹批量导入 |
| 输出 | 分子属性 Excel + 3D 结构文件 + 相似性矩阵 Excel |
| 亮点 | 10 种官能团自动识别 · ETKDGv3 3D 生成 + MMFF/UFF 力场优化 · 4 种输出格式（SDF/MOL/PDB/MOL2）· Morgan 指纹 + Tanimoto 相似性 |

---

## 注意事项

1. **DisGeNET_疾病靶点检索.py** 首次运行需申请 DisGeNET 免费学术 API Key（脚本内置申请引导，审核通常 2-7 天）
2. **SuperPred_靶点预测.py** 和 **SwissTarget_靶点预测.py** 依赖国外网站在线可用，建议在网络稳定环境下使用；内置防封机制可一定程度规避限流
3. **小分子结构处理工具.py** 的 MOL2 格式输出依赖 OpenBabel（`pip install openbabel-wheel`），若未安装会自动回退到 SDF
4. 批量预测类脚本（5/6/7/8）建议按推荐间隔运行，避免对服务器造成压力

---

## 项目信息

- **作者**: shenjianlin
- **仓库**: `git@github.com:sjl-openmywork/Network-toxicology-script.git`