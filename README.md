# 网络毒理学靶点预测脚本套件

一站式网络毒理学/药理学研究工具集，覆盖 5 大权威数据库靶点查询 + 3 大 AI 平台靶点预测 + 分子属性计算与结构处理。

终端交互式操作，零编程基础上手，结果一键导出 Excel。

---

## 脚本清单

| # | 脚本 | 功能 | 分类 |
|---|------|------|------|
| 1 | `TTP.py` | TTD 疾病靶点搜索 | 数据库查询 |
| 2 | `chembl_target_search_interactive.py` | ChEMBL 化合物靶点检索 | 数据库查询 |
| 3 | `disgenet.py` | DisGeNET 疾病靶点检索 | 数据库查询 |
| 4 | `open_targets.py` | Open Targets 疾病靶点检索 | 数据库查询 |
| 5 | `pharmgkb.py` | PharmGKB 基因靶点搜索 | 数据库查询 |
| 6 | `superpred.py` | SuperPred AI 靶点预测 | AI 预测 |
| 7 | `swiss_target_prediction_gui.py` | SwissTargetPrediction 靶点预测 | AI 预测 |
| 8 | `targetnet_interactive.py` | TargetNet QSAR 靶点预测 | AI 预测 |
| 9 | `计算化学_小分子处理.py` | 分子属性计算 & 结构处理 | 本地计算 |

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
python TTP.py
python chembl_target_search_interactive.py
python disgenet.py
```

所有脚本均采用终端交互菜单，按提示操作即可。

---

## 脚本详情

### 一、数据库查询类（5 款）

#### 1. TTD 疾病靶点搜索 — `TTP.py`

| 项目 | 说明 |
|------|------|
| 数据源 | TTD (Therapeutic Target Database) — 全球最大的治疗靶点数据库 |
| 输入 | 疾病英文名称（如 Alzheimer、Diabetes）或 ICD-11 编码 |
| 输出 | Excel (.xlsx)，含靶点 ID、基因名、相关疾病、代表药物 |
| 亮点 | 基本模式（秒级）与详细模式（含 UniProt/染色体位置/功能描述） |

#### 2. ChEMBL 化合物靶点检索 — `chembl_target_search_interactive.py`

| 项目 | 说明 |
|------|------|
| 数据源 | ChEMBL (EMBL-EBI) — 全球最大开源生物活性数据库 |
| 输入 | SMILES 结构式 或 化合物名称（自动识别） |
| 输出 | 合并 Excel + 每个化合物单独 CSV |
| 亮点 | 3 种搜索模式（精确/相似/子结构）· 8 线程并发获取基因详情 · 可设活性阈值与物种过滤 · Excel/CSV 批量输入 + GUI 弹窗选列 |

#### 3. DisGeNET 疾病靶点检索 — `disgenet.py`

| 项目 | 说明 |
|------|------|
| 数据源 | DisGeNET — 全球最大的基因-疾病关联知识平台 |
| 输入 | 疾病英文名称 或 疾病 ID（UMLS/MONDO/OMIM/MeSH 等） |
| 输出 | Excel (.xlsx)，含 3 个 Sheet（基因疾病关联/TOP100/汇总统计） |
| 亮点 | 专业评分体系（score/DSI/DPI/pLI/EI）· 可下载全部疾病列表（≤10,000 条）· 独有 Disease Ontology & MONDO 本体下载 |
| 注意 | 需申请 DisGeNET 免费学术 API Key（脚本内置申请引导） |

#### 4. Open Targets 疾病靶点检索 — `open_targets.py`

| 项目 | 说明 |
|------|------|
| 数据源 | Open Targets Platform (EMBL-EBI + GSK + Sanger 等联合项目) |
| 输入 | 疾病名称（支持中英文）或疾病 ID（MONDO/EFO） |
| 输出 | Excel (.xlsx)，含 3 个 Sheet，每种证据类型独立成列 |
| 亮点 | 8 维度证据评分（遗传关联/体细胞突变/RNA 表达/通路/已知药物/动物模型/文献/临床）· 多疾病批量查询 · 12 种常见疾病快捷入口 |

#### 5. PharmGKB 基因靶点搜索 — `pharmgkb.py`

| 项目 | 说明 |
|------|------|
| 数据源 | PharmGKB (NIH 资助) — 药物基因组学权威数据库 |
| 输入 | 疾病或药物关键词（如 Diabetes、Aspirin） |
| 输出 | Excel (.xlsx)，基因符号列表 |
| 亮点 | 自动读取总数 · 智能翻页 · 请求重试与速率控制 · 去重排序 |

### 二、AI 靶点预测类（3 款）

通过浏览器自动化，将化合物 SMILES 提交至 AI 预测平台，自动解析结果。

#### 6. SuperPred 靶点预测 — `superpred.py`

| 项目 | 说明 |
|------|------|
| 平台 | SuperPred (德国 Charité 医学院) — 基于机器学习的靶点预测 |
| 输入 | 化合物 SMILES（支持单化合物/Excel 批量/CSV 批量） |
| 输出 | 每个化合物单独 CSV + 合并 Excel + 汇总报告 + 失败列表 |
| 亮点 | Playwright / Selenium 双引擎 · GUI 弹窗选列 · JS API + DOM 双路径结果提取 · 防封机制 · 5 个内置化合物快速体验 |

#### 7. SwissTargetPrediction 自动化预测 — `swiss_target_prediction_gui.py`

| 项目 | 说明 |
|------|------|
| 平台 | SwissTargetPrediction (瑞士生物信息学研究所 SIB) — 高精度靶点预测 |
| 输入 | Excel (.xlsx) 或 CSV 文件，含 SMILES 列 |
| 输出 | 每个化合物单独 CSV + 合并 Excel + 汇总报告 |
| 亮点 | 7 步交互式配置向导 · RDKit SMILES 合法性预检 · 人/小鼠/大鼠 3 物种 · 强大防封机制 |

#### 8. TargetNet QSAR 靶点预测 — `targetnet_interactive.py`

| 项目 | 说明 |
|------|------|
| 平台 | TargetNet (中南大学) — 基于 QSAR 模型的靶点预测 |
| 输入 | SMILES 字符串 或 .smi 文件 或 Excel 文件 |
| 输出 | 每个化合物单独 CSV + 合并 Excel（csv/xlsx/txt 可选） |
| 亮点 | 7 种分子指纹可选 · 4 种筛选指标 + 可调阈值 · 快速/手动双模式 · 智能自动分批 · GUI 弹窗选列 |

### 三、分子处理类（1 款）

#### 9. 计算化学·小分子处理工具 — `计算化学_小分子处理.py`

| 项目 | 说明 |
|------|------|
| 依赖 | 纯本地计算（RDKit + OpenBabel），无需网络 |
| 输入 | Excel (.xlsx) / SDF / MOL / MOL2 / PDB，支持文件夹批量导入 |
| 输出 | 分子属性 Excel + 3D 结构文件 + 相似性矩阵 Excel |
| 亮点 | 10 种官能团自动识别 · ETKDGv3 3D 生成 + MMFF/UFF 力场优化 · 4 种输出格式（SDF/MOL/PDB/MOL2）· Morgan 指纹 + Tanimoto 相似性 |

---

## 注意事项

1. **disgenet.py** 首次运行需申请 DisGeNET 免费学术 API Key（脚本内置引导流程，审核通常 2-7 天）
2. **superpred.py** 和 **swiss_target_prediction_gui.py** 依赖国外网站在线可用，建议在网络稳定环境下使用；内置防封机制可一定程度规避限流
3. **计算化学_小分子处理.py** 的 MOL2 格式输出依赖 OpenBabel（`pip install openbabel-wheel`），若未安装会自动回退到 SDF
4. 批量预测类脚本（2/6/7/8）建议按推荐间隔运行，避免对服务器造成压力

---

## 项目信息

- **作者**: shenjianlin
- **仓库**: `git@github.com:sjl-openmywork/Network-toxicology-script.git`