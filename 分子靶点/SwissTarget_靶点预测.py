# -*- coding: utf-8 -*-
"""
SwissTargetPrediction 交互式自动化化合物靶点预测
==================================================
功能：
  1. 交互式配置所有运行参数（文件、物种、间隔等）
  2. 用户手动指定 SMILES 列和名称列
  3. SMILES 合法性预检（支持 RDKit 精确验证 / 正则兜底）
  4. 增强防封 IP 机制（UA 轮换、随机视口、批次冷却、指数退避）
  5. 每个化合物生成单独的 Excel 文件 + 合并汇总文件

使用方法：
  python SwissTarget_靶点预测.py

依赖安装：
  pip install playwright openpyxl pandas
  pip install rdkit   (可选，用于精确 SMILES 验证)
"""

import os
import re
import sys
import time
import random
import csv
import logging
import traceback
from pathlib import Path
from datetime import datetime

import pandas as pd
import openpyxl
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

# ============================================================
# 终端颜色支持（Windows 10+）
# ============================================================
if sys.platform == "win32":
    os.system("")  # 启用 ANSI 支持


class Color:
    """终端 ANSI 颜色"""
    RESET = "\033[0m"
    BOLD = "\033[1m"
    DIM = "\033[2m"
    RED = "\033[91m"
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    BLUE = "\033[94m"
    MAGENTA = "\033[95m"
    CYAN = "\033[96m"
    WHITE = "\033[97m"


def cprint(text, color="", bold=False, end="\n"):
    """彩色打印"""
    prefix = ""
    if bold:
        prefix += Color.BOLD
    if color:
        prefix += color
    print(f"{prefix}{text}{Color.RESET}", end=end)


def print_banner():
    """打印横幅"""
    cprint("╔" + "═" * 58 + "╗", Color.CYAN, bold=True)
    cprint("║" + "  SwissTargetPrediction 交互式化合物靶点预测".center(52) + "║", Color.CYAN, bold=True)
    cprint("║" + "  author: shenjianlin".center(50) + "║", Color.CYAN)
    cprint("║" + "  site: git@github.com:sjl-openmywork/Network-toxicology-script.git".center(38) + "║", Color.CYAN)
    cprint("╚" + "═" * 58 + "╝", Color.CYAN, bold=True)
    print()


def print_section(title):
    """打印章节标题"""
    cprint(f"▸ {title}", Color.CYAN, bold=True)


def print_ok(text):
    """打印成功信息"""
    cprint(f"  ✓ {text}", Color.GREEN)


def print_warn(text):
    """打印警告信息"""
    cprint(f"  ⚠ {text}", Color.YELLOW)


def print_err(text):
    """打印错误信息"""
    cprint(f"  ✗ {text}", Color.RED)


def print_info(text):
    """打印普通信息"""
    cprint(f"  {text}", Color.DIM)


# ============================================================
# 运行配置（运行时由用户交互设定）
# ============================================================

class RunConfig:
    """运行配置容器"""
    def __init__(self):
        self.input_file = ""
        self.species = "Homo_sapiens"
        self.output_dir = "swisstarget_results"
        self.request_interval = 5
        self.random_delay_range = (1, 3)
        self.page_load_timeout = 120000
        self.headless = True
        self.probability_threshold = 0.0
        self.browser_channel = "chrome"
        # 列映射
        self.smiles_col = "SMILES"
        self.name_col = None  # None 表示自动生成编号
        # 防封配置
        self.batch_size = 20          # 每批处理的化合物数量
        self.batch_cooldown = (30, 60)  # 批次间冷却时间（秒）
        self.max_retries = 3          # 失败最大重试次数
        self.use_random_ua = True     # 随机 User-Agent
        self.use_random_viewport = True  # 随机视口大小
        self.use_mouse_simulation = True  # 模拟鼠标移动


# ============================================================
# SMILES 验证
# ============================================================

# 尝试导入 RDKit
try:
    from rdkit import Chem
    HAS_RDKIT = True
except ImportError:
    HAS_RDKIT = False

# 基础 SMILES 正则（兜底验证）
# 匹配常见 SMILES 特征：原子、括号、键符号、环编号等
SMILES_REGEX = re.compile(
    r'^['
    r'CNOSPFIBrcnosp'           # 常见原子（大小写）
    r'\[\]'                      # 方括号（金属/同位素等）
    r'\(\)'                      # 圆括号（分支）
    r'0-9'                       # 环编号
    r'=#\-+\\/'                  # 键类型
    r'\.@/'                      # 立体化学/芳香性
    r'Cl|Br|Si|Se|Te|Ge|Sn|As|Hg|Pb|Bi'  # 双字符原子（部分）
    r']+$'
)

# 更严格的正则：至少包含一个碳原子或常见有机原子
SMILES_STRICT_REGEX = re.compile(
    r'[CNOSPFIBcnos]'  # 至少包含一个有机原子
)


def validate_smiles(smiles_str):
    """
    验证单个 SMILES 字符串的合法性。

    返回：
        (is_valid: bool, reason: str)
    """
    if not smiles_str or not isinstance(smiles_str, str):
        return False, "空值或非字符串"

    s = smiles_str.strip()
    if not s:
        return False, "空字符串"

    # 基本长度检查
    if len(s) < 2:
        return False, f"过短（{len(s)}字符）"
    if len(s) > 5000:
        return False, f"过长（{len(s)}字符）"

    # 使用 RDKit 精确验证（如果可用）
    if HAS_RDKIT:
        mol = Chem.MolFromSmiles(s)
        if mol is None:
            return False, "RDKit 解析失败"
        return True, "RDKit 验证通过"

    # 兜底：正则验证
    if not SMILES_STRICT_REGEX.search(s):
        return False, "不含常见有机原子"

    # 括号匹配检查
    if s.count('(') != s.count(')'):
        return False, "括号不匹配"
    if s.count('[') != s.count(']'):
        return False, "方括号不匹配"

    # 非法字符检查（允许常见 SMILES 字符）
    allowed_chars = set(
        'CNOSPFIBrcnosp0123456789'
        '()[]{}'
        '=#\\-+/:.@'
        'lreagiumdvtkwy'  # 双字符原子的组成部分
    )
    invalid_chars = set(s) - allowed_chars
    if invalid_chars:
        return False, f"含非法字符: {''.join(sorted(invalid_chars))}"

    return True, "正则验证通过"


def validate_smiles_batch(df, smiles_col):
    """
    批量验证 SMILES。

    返回：
        (valid_df: DataFrame, invalid_list: list[dict])
    """
    valid_indices = []
    invalid_list = []

    for idx, row in df.iterrows():
        s = str(row[smiles_col]).strip()
        is_valid, reason = validate_smiles(s)
        if is_valid:
            valid_indices.append(idx)
        else:
            invalid_list.append({
                "行号": idx + 1,
                "SMILES": s[:80] if s else "(空)",
                "原因": reason,
            })

    valid_df = df.loc[valid_indices].copy()
    return valid_df, invalid_list


# ============================================================
# 日志配置
# ============================================================

def setup_logging(output_dir):
    """配置日志（同时输出到控制台和文件）"""
    os.makedirs(output_dir, exist_ok=True)

    log_file = os.path.join(output_dir, f"swisstarget_{datetime.now().strftime('%Y%m%d')}.log")

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    logger = logging.getLogger("SwissTargetPrediction")
    # 降低控制台日志级别
    for h in logger.handlers:
        if isinstance(h, logging.StreamHandler) and h.stream == sys.stdout:
            h.setLevel(logging.WARNING)  # 控制台只显示 WARNING+
    return logger


logger = logging.getLogger("SwissTargetPrediction")


# ============================================================
# 防封 IP 辅助工具
# ============================================================

# 真实浏览器 User-Agent 列表（定期更新）
USER_AGENTS = [
    # Chrome on Windows
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    # Chrome on Mac
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    # Edge on Windows
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36 Edg/125.0.0.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36 Edg/124.0.0.0",
    # Firefox on Windows
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:126.0) Gecko/20100101 Firefox/126.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
    # Firefox on Mac
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:126.0) Gecko/20100101 Firefox/126.0",
]

# 常见视口分辨率
VIEWPORT_SIZES = [
    {"width": 1920, "height": 1080},
    {"width": 1536, "height": 864},
    {"width": 1440, "height": 900},
    {"width": 1366, "height": 768},
    {"width": 1280, "height": 720},
    {"width": 1600, "height": 900},
    {"width": 1680, "height": 1050},
    {"width": 2560, "height": 1440},
]


def get_random_ua():
    """获取随机 User-Agent"""
    return random.choice(USER_AGENTS)


def get_random_viewport():
    """获取随机视口大小"""
    return random.choice(VIEWPORT_SIZES)


def random_sleep(min_sec, max_sec):
    """随机等待，带对数正态分布"""
    mean = (min_sec + max_sec) / 2
    sigma = (max_sec - min_sec) / 4
    delay = random.lognormvariate(mean, sigma)
    delay = max(min_sec, min(max_sec * 1.5, delay))
    return delay


# ============================================================
# 交互式配置向导
# ============================================================


def interactive_setup():
    """交互式配置运行参数，返回 RunConfig 对象"""
    cfg = RunConfig()

    print_banner()
    cprint("欢迎使用 SwissTargetPrediction 交互式预测工具！", Color.WHITE, bold=True)
    print("本工具将引导您完成所有参数设置，请按提示输入。\n")
    cprint("提示：直接按 Enter 将使用括号 [ ] 中的默认值。\n", Color.DIM)

    # ── 步骤 1：输入文件 ──
    print_section("步骤 1/7：选择输入文件")
    print("  支持 .xlsx / .xls / .csv 格式\n")

    # 扫描当前目录的 Excel/CSV 文件
    current_dir = os.getcwd()
    available_files = []
    for ext in [".xlsx", ".xls", ".csv"]:
        available_files.extend(Path(current_dir).glob(f"*{ext}"))

    if available_files:
        print("  当前目录下找到以下文件：")
        for i, f in enumerate(available_files, 1):
            print(f"    [{i}] {f.name}")
        print("    [0] 手动输入路径\n")
        choice = input("  请选择文件编号 [0]: ").strip()
        if choice and choice != "0":
            try:
                idx = int(choice) - 1
                if 0 <= idx < len(available_files):
                    cfg.input_file = str(available_files[idx])
                else:
                    print_warn("编号无效，请输入路径")
                    cfg.input_file = ""
            except ValueError:
                print_warn("输入无效，请输入路径")
                cfg.input_file = ""

    if not cfg.input_file:
        default_input = "compounds.xlsx"
        inp = input(f"  请输入输入文件路径 [{default_input}]: ").strip()
        cfg.input_file = inp if inp else default_input

    # 验证文件并读取原始数据
    raw_df = None
    while True:
        try:
            raw_df = _read_raw_file(cfg.input_file)
            print_ok(f"成功读取文件：{len(raw_df)} 行，{len(raw_df.columns)} 列")
            break
        except FileNotFoundError:
            print_err(f"文件不存在: '{cfg.input_file}'")
            inp = input("  请重新输入文件路径（或按 Enter 退出）: ").strip()
            if not inp:
                print_warn("未指定输入文件，程序退出")
                sys.exit(0)
            cfg.input_file = inp
        except ValueError as e:
            print_err(f"文件格式错误: {e}")
            inp = input("  请重新输入文件路径（或按 Enter 退出）: ").strip()
            if not inp:
                print_warn("未指定输入文件，程序退出")
                sys.exit(0)
            cfg.input_file = inp
    print()

    # ── 步骤 2：指定列 ──
    print_section("步骤 2/7：指定数据列")
    print("  文件包含以下列：\n")
    columns = list(raw_df.columns)
    for i, col in enumerate(columns, 1):
        # 显示列名和前几个值作为预览
        sample_vals = raw_df[col].dropna().head(2).astype(str).tolist()
        sample_str = " | ".join(v[:30] + "..." if len(v) > 30 else v for v in sample_vals)
        print(f"    [{i}] {col}  →  {sample_str}")
    print()

    # 选择 SMILES 列
    smiles_col_idx = None
    # 尝试自动推荐
    auto_smiles = None
    for i, col in enumerate(columns, 1):
        if col.lower() in ("smiles", "smi", "smiles_string", "canonical_smiles"):
            auto_smiles = i
            break
    if auto_smiles is None:
        for i, col in enumerate(columns, 1):
            if "smiles" in col.lower() or "smi" in col.lower():
                auto_smiles = i
                break

    default_smiles = str(auto_smiles) if auto_smiles else "1"
    while True:
        inp = input(f"  请选择 SMILES 列的编号 [{default_smiles}]: ").strip()
        inp = inp if inp else default_smiles
        try:
            idx = int(inp) - 1
            if 0 <= idx < len(columns):
                cfg.smiles_col = columns[idx]
                print_ok(f"SMILES 列: '{cfg.smiles_col}'")
                smiles_col_idx = idx
                break
            else:
                print_err("编号超出范围，请重新选择")
        except ValueError:
            print_err("请输入有效数字")

    # 选择名称列（可选）
    print()
    auto_name = None
    for i, col in enumerate(columns, 1):
        if i - 1 == smiles_col_idx:
            continue
        if col.lower() in ("compound_name", "name", "compound", "化合物名称", "名称", "compound_id", "id"):
            auto_name = i
            break
    if auto_name is None:
        for i, col in enumerate(columns, 1):
            if i - 1 == smiles_col_idx:
                continue
            if any(kw in col.lower() for kw in ["name", "名称", "compound", "title"]):
                auto_name = i
                break

    default_name = str(auto_name) if auto_name else "0"
    print(f"\n  名称列可选（用于为化合物命名），选 0 表示自动生成编号")
    while True:
        inp = input(f"  请选择名称列的编号 [{default_name}]（0=自动生成）: ").strip()
        inp = inp if inp else default_name
        try:
            idx = int(inp)
            if idx == 0:
                cfg.name_col = None
                print_ok("名称列: 自动生成编号")
                break
            elif 1 <= idx <= len(columns):
                if idx - 1 == smiles_col_idx:
                    print_err("名称列不能与 SMILES 列相同，请重新选择")
                    continue
                cfg.name_col = columns[idx - 1]
                print_ok(f"名称列: '{cfg.name_col}'")
                break
            else:
                print_err("编号超出范围，请重新选择")
        except ValueError:
            print_err("请输入有效数字")
    print()

    # ── 步骤 3：物种选择 ──
    print_section("步骤 3/7：选择物种")
    print("  [1] Homo sapiens（人类）- 推荐用于网络药理学研究")
    print("  [2] Mus musculus（小鼠）")
    print("  [3] Rattus norvegicus（大鼠）\n")
    species_map = {"1": "Homo_sapiens", "2": "Mus_musculus", "3": "Rattus_norvegicus"}
    species_names = {
        "Homo_sapiens": "Homo sapiens（人类）",
        "Mus_musculus": "Mus musculus（小鼠）",
        "Rattus_norvegicus": "Rattus norvegicus（大鼠）",
    }
    choice = input("  请选择物种 [1]: ").strip()
    cfg.species = species_map.get(choice, "Homo_sapiens")
    print_ok(f"已选择: {species_names[cfg.species]}")
    print()

    # ── 步骤 4：输出目录 ──
    print_section("步骤 4/7：设置输出目录")
    inp = input(f"  请输入输出目录路径 [swisstarget_results]: ").strip()
    cfg.output_dir = inp if inp else "swisstarget_results"
    print_ok(f"输出目录: {cfg.output_dir}")
    print()

    # ── 步骤 5：请求间隔 ──
    print_section("步骤 5/7：设置请求间隔")
    print("  为避免对服务器造成压力，建议 5-15 秒\n")
    inp = input("  请输入请求间隔秒数 [5]: ").strip()
    try:
        cfg.request_interval = max(1, min(60, int(inp))) if inp else 5
    except ValueError:
        cfg.request_interval = 5
    print_ok(f"请求间隔: {cfg.request_interval} 秒")
    print()

    # ── 步骤 6：高级选项 ──
    print_section("步骤 6/7：高级选项")
    print("  [1] 使用默认设置（推荐）")
    print("  [2] 自定义设置\n")
    choice = input("  请选择 [1]: ").strip()

    if choice == "2":
        # 无头模式
        print()
        cprint("  浏览器模式:", Color.WHITE, bold=True)
        print("    [1] 无头模式 - 后台静默运行（推荐）")
        print("    [2] 可见模式 - 显示浏览器窗口，可观察运行过程")
        hm = input("  请选择 [1]: ").strip()
        cfg.headless = (hm != "2")
        print_ok(f"浏览器模式: {'无头' if cfg.headless else '可见'}")

        # 概率阈值
        cprint("  概率阈值:", Color.WHITE, bold=True)
        print("    仅保留概率 >= 设定值的靶点，0 表示保留全部")
        inp = input("  请输入概率阈值 [0]: ").strip()
        try:
            cfg.probability_threshold = max(0.0, min(1.0, float(inp))) if inp else 0.0
        except ValueError:
            cfg.probability_threshold = 0.0
        print_ok(f"概率阈值: {cfg.probability_threshold}")

        # 随机延迟
        cprint("  随机延迟范围:", Color.WHITE, bold=True)
        print("    在基础间隔上额外增加随机延迟（下限, 上限）")
        inp = input("  请输入随机延迟下限秒数 [1]: ").strip()
        try:
            low = float(inp) if inp else 1.0
        except ValueError:
            low = 1.0
        inp = input("  请输入随机延迟上限秒数 [3]: ").strip()
        try:
            high = float(inp) if inp else 3.0
        except ValueError:
            high = 3.0
        cfg.random_delay_range = (min(low, high), max(low, high))
        print_ok(f"随机延迟范围: {cfg.random_delay_range}")

        # 批次冷却
        print()
        cprint("  批次冷却（防封增强）:", Color.WHITE, bold=True)
        print("    每处理 N 个化合物后暂停较长时间，模拟人类行为")
        inp = input("  每批处理化合物数量 [20]: ").strip()
        try:
            cfg.batch_size = max(5, min(100, int(inp))) if inp else 20
        except ValueError:
            cfg.batch_size = 20
        inp = input(f"  批次冷却时间下限秒数 [30]: ").strip()
        try:
            cool_low = int(inp) if inp else 30
        except ValueError:
            cool_low = 30
        inp = input(f"  批次冷却时间上限秒数 [60]: ").strip()
        try:
            cool_high = int(inp) if inp else 60
        except ValueError:
            cool_high = 60
        cfg.batch_cooldown = (min(cool_low, cool_high), max(cool_low, cool_high))
        print_ok(f"批次大小: {cfg.batch_size}，冷却: {cfg.batch_cooldown[0]}-{cfg.batch_cooldown[1]}秒")

        # 重试次数
        print()
        cprint("  失败重试:", Color.WHITE, bold=True)
        inp = input("  单个化合物预测失败时最大重试次数 [3]: ").strip()
        try:
            cfg.max_retries = max(0, min(10, int(inp))) if inp else 3
        except ValueError:
            cfg.max_retries = 3
        print_ok(f"最大重试次数: {cfg.max_retries}")
    else:
        print_ok("使用默认高级设置")
    print()

    # ── 步骤 7：确认 ──
    print_section("步骤 7/7：确认设置")
    print()
    cprint("  ──────────── 运行设置预览 ────────────", Color.BOLD)
    print(f"  输入文件:        {cfg.input_file}")
    print(f"  SMILES 列:       {cfg.smiles_col}")
    print(f"  名称列:          {cfg.name_col or '自动生成编号'}")
    print(f"  物种:            {species_names[cfg.species]}")
    print(f"  输出目录:        {cfg.output_dir}")
    print(f"  请求间隔:        {cfg.request_interval} 秒（+{cfg.random_delay_range[0]}-{cfg.random_delay_range[1]} 秒随机）")
    print(f"  批次冷却:        每 {cfg.batch_size} 个暂停 {cfg.batch_cooldown[0]}-{cfg.batch_cooldown[1]} 秒")
    print(f"  失败重试:        最多 {cfg.max_retries} 次")
    print(f"  浏览器模式:      {'无头（后台运行）' if cfg.headless else '可见（显示窗口）'}")
    print(f"  概率阈值:        {cfg.probability_threshold}" if cfg.probability_threshold > 0 else f"  概率阈值:        无（保留全部结果）")
    cprint("  ──────────────────────────────────────", Color.BOLD)
    print()

    confirm = input("  确认开始运行？[Y/n]: ").strip().lower()
    if confirm == "n":
        print_warn("已取消运行")
        sys.exit(0)

    print()
    return cfg, raw_df


def _read_raw_file(filepath):
    """读取原始文件，返回 DataFrame（不做列过滤）"""
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"文件不存在: {filepath}")

    if filepath.endswith((".xlsx", ".xls")):
        df = pd.read_excel(filepath)
    elif filepath.endswith(".csv"):
        df = pd.read_csv(filepath)
    else:
        raise ValueError("不支持的文件格式，请使用 .xlsx 或 .csv 文件")

    if len(df) == 0:
        raise ValueError("文件为空")

    return df


# ============================================================
# 进度显示
# ============================================================


def print_progress(current, total, compound_name, status, elapsed=0):
    """打印进度条"""
    bar_width = 30
    pct = current / total if total > 0 else 0
    filled = int(bar_width * pct)
    bar = "█" * filled + "░" * (bar_width - filled)
    elapsed_str = f"{elapsed / 60:.1f}min" if elapsed > 60 else f"{elapsed:.0f}s"

    line = f"\r  [{bar}] {current}/{total} ({pct * 100:.0f}%) | {compound_name[:20]:<20} | {status:<12} | {elapsed_str}"
    sys.stdout.write(line + " " * 10)
    sys.stdout.flush()


# ============================================================
# 核心预测器
# ============================================================


class SwissTargetPredictor:
    """SwissTargetPrediction 自动化预测器（基于 Playwright）"""

    BASE_URL = "https://www.swisstargetprediction.ch/"

    def __init__(self, config):
        self.cfg = config
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None
        self._request_count = 0  # 请求计数（用于批次冷却）

    def _init_browser(self):
        """初始化浏览器（带防封措施）"""
        self.playwright = sync_playwright().start()

        launch_args = [
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--disable-blink-features=AutomationControlled",
            "--disable-infobars",
            "--disable-extensions",
            "--disable-popup-blocking",
            "--disable-translate",
            "--metrics-recording-only",
            "--no-first-run",
            "--safebrowsing-disable-automatic-fetching",
        ]

        browser_kwargs = {
            "headless": self.cfg.headless,
            "args": launch_args,
        }
        if self.cfg.browser_channel:
            browser_kwargs["channel"] = self.cfg.browser_channel

        self.browser = self.playwright.chromium.launch(**browser_kwargs)

        # 随机 User-Agent 和视口
        ua = get_random_ua() if self.cfg.use_random_ua else USER_AGENTS[0]
        viewport = get_random_viewport() if self.cfg.use_random_viewport else {"width": 1920, "height": 1080}

        self.context = self.browser.new_context(
            viewport=viewport,
            user_agent=ua,
            locale="en-US",
            timezone_id="America/New_York",
            color_scheme="light",
            java_script_enabled=True,
            ignore_https_errors=True,
        )

        # 隐藏 webdriver 特征（增强版）
        self.context.add_init_script("""
            // 隐藏 webdriver 属性
            Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
            
            // 伪造 plugins
            Object.defineProperty(navigator, 'plugins', {
                get: () => [1, 2, 3, 4, 5],
            });
            
            // 伪造 languages
            Object.defineProperty(navigator, 'languages', {
                get: () => ['en-US', 'en'],
            });
            
            // 隐藏自动化标志
            window.chrome = { runtime: {} };
            
            // 伪造 permissions
            const originalQuery = window.navigator.permissions.query;
            window.navigator.permissions.query = (parameters) => (
                parameters.name === 'notifications' ?
                    Promise.resolve({ state: Notification.permission }) :
                    originalQuery(parameters)
            );
        """)

        self.page = self.context.new_page()
        self.page.set_default_timeout(self.cfg.page_load_timeout)
        logger.info(f"浏览器初始化成功 | UA: {ua[:50]}... | 视口: {viewport}")

    def _close_browser(self):
        """关闭浏览器"""
        for obj in [self.context, self.browser, self.playwright]:
            if obj:
                try:
                    if hasattr(obj, 'close'):
                        obj.close()
                    elif hasattr(obj, 'stop'):
                        obj.stop()
                except Exception:
                    pass
        self.context = None
        self.browser = None
        self.playwright = None
        self.page = None
        logger.info("浏览器已关闭")

    def _simulate_human_behavior(self):
        """模拟人类行为（鼠标移动、滚动等）"""
        if not self.cfg.use_mouse_simulation:
            return
        try:
            # 随机鼠标移动
            x = random.randint(100, 800)
            y = random.randint(100, 600)
            self.page.mouse.move(x, y, steps=random.randint(5, 15))
            self.page.wait_for_timeout(random.randint(100, 300))

            # 偶尔滚动
            if random.random() < 0.3:
                self.page.mouse.wheel(0, random.randint(-100, 100))
                self.page.wait_for_timeout(random.randint(100, 200))
        except Exception:
            pass  # 鼠标模拟失败不影响主流程

    def _batch_cooldown_check(self):
        """检查是否需要批次冷却"""
        if self._request_count > 0 and self._request_count % self.cfg.batch_size == 0:
            cooldown = random.uniform(*self.cfg.batch_cooldown)
            logger.info(f"批次冷却: 已处理 {self._request_count} 个，暂停 {cooldown:.0f} 秒")
            print_info(f"\n  ⏸ 批次冷却: 暂停 {cooldown:.0f} 秒（已处理 {self._request_count} 个）...")
            time.sleep(cooldown)
            # 冷却后模拟人类行为
            self._simulate_human_behavior()

    def predict_single(self, smiles, compound_name="Unknown", retry=0):
        """对单个化合物进行靶点预测（带重试）"""
        if not self.page:
            self._init_browser()

        try:
            logger.info(f"开始预测: {compound_name} | SMILES: {smiles[:50]}...")

            # 模拟人类行为
            self._simulate_human_behavior()

            # 1. 访问首页
            self.page.goto(self.BASE_URL, wait_until="domcontentloaded")
            self.page.wait_for_timeout(random.randint(800, 1500))

            # 2. 选择物种
            self.page.click(
                f'input[type="radio"][name="organism"][value="{self.cfg.species}"]'
            )
            self.page.wait_for_timeout(random.randint(300, 600))

            # 3. 提交表单（target 改为 _self 避免弹窗）
            self.page.evaluate(
                """(smiles) => {
                    const form = document.getElementById('myForm') || document.forms[0];
                    if (form) {
                        form.target = '_self';
                        form.smiles.value = smiles;
                        form.ioi.value = 2;
                        form.organism.value = document.querySelector(
                            'input[type="radio"][name="organism"]:checked'
                        ).value;
                    }
                    const box = document.getElementById('smilesBox');
                    box.value = smiles;
                    formSubmit();
                }""",
                smiles,
            )

            # 4. 等待结果页
            self.page.wait_for_url('**/result.php**', timeout=self.cfg.page_load_timeout)
            self.page.wait_for_selector('#resultTable', timeout=self.cfg.page_load_timeout)
            self.page.wait_for_selector('#resultTable tbody tr', timeout=60000)
            self.page.wait_for_timeout(random.randint(1500, 2500))

            # 5. 显示所有结果
            try:
                length_select = self.page.query_selector('select[name="resultTable_length"]')
                if length_select:
                    length_select.select_option(value="-1")
                    self.page.wait_for_timeout(random.randint(1500, 2500))
            except Exception:
                pass

            # 6. 解析表格
            results = self._parse_result_table(self.page, compound_name, smiles)

            self._request_count += 1
            logger.info(f"预测完成: {compound_name} | 获取到 {len(results)} 个靶点")
            return results

        except PlaywrightTimeout:
            logger.error(f"预测超时: {compound_name}")
            if retry < self.cfg.max_retries:
                return self._retry_with_backoff(smiles, compound_name, retry, "超时")
            return None
        except Exception as e:
            logger.error(f"预测失败: {compound_name} | 错误: {e}")
            if retry < self.cfg.max_retries:
                return self._retry_with_backoff(smiles, compound_name, retry, str(e))
            return None

    def _retry_with_backoff(self, smiles, compound_name, retry, error_msg):
        """指数退避重试"""
        retry += 1
        backoff = min(30, 2 ** retry + random.uniform(0, 2))
        logger.info(f"重试 {retry}/{self.cfg.max_retries}: {compound_name} | 等待 {backoff:.1f} 秒 | 原因: {error_msg}")
        print_warn(f"重试 {retry}/{self.cfg.max_retries}（等待 {backoff:.0f}s）")
        time.sleep(backoff)

        # 重试前刷新浏览器上下文
        try:
            if self.context:
                self.context.clear_cookies()
        except Exception:
            pass

        return self.predict_single(smiles, compound_name, retry)

    def _parse_result_table(self, page, compound_name, smiles):
        """解析结果表格"""
        results = []

        js_code = """
        () => {
            const rows = document.querySelectorAll('#resultTable tbody tr');
            const data = [];
            rows.forEach(row => {
                const cells = row.querySelectorAll('td');
                if (cells.length < 7) return;
                try {
                    const probSpan = cells[5].querySelector('span');
                    const probability = probSpan ? parseFloat(probSpan.textContent.trim()) : 0;
                    const activesLinks = cells[6].querySelectorAll('a');
                    const known3d = activesLinks.length >= 1 ? parseInt(activesLinks[0].textContent.trim()) : 0;
                    const known2d = activesLinks.length >= 2 ? parseInt(activesLinks[1].textContent.trim()) : 0;
                    data.push({
                        target: cells[0].textContent.trim(),
                        common_name: cells[1].textContent.trim(),
                        uniprot_id: cells[2].textContent.trim(),
                        chembl_id: cells[3].textContent.trim(),
                        target_class: cells[4].textContent.trim(),
                        probability: probability,
                        known_3d: known3d,
                        known_2d: known2d
                    });
                } catch(e) {}
            });
            return data;
        }
        """

        try:
            raw_data = page.evaluate(js_code)
            for item in raw_data:
                if self.cfg.probability_threshold > 0 and item["probability"] < self.cfg.probability_threshold:
                    continue
                result = {
                    "Compound_Name": compound_name,
                    "SMILES": smiles,
                    "Target": item["target"],
                    "Common_Name": item["common_name"],
                    "Uniprot_ID": item["uniprot_id"],
                    "ChEMBL_ID": item["chembl_id"],
                    "Target_Class": item["target_class"],
                    "Probability": round(item["probability"], 6),
                    "Known_Actives_3D": item["known_3d"],
                    "Known_Actives_2D": item["known_2d"],
                }
                results.append(result)
        except Exception as e:
            logger.warning(f"解析表格出错: {e}")

        return results

    def batch_predict(self, compounds_df, smiles_col, name_col=None):
        """批量预测"""
        all_results = {}
        total = len(compounds_df)
        start_time = time.time()

        cprint("\n" + "─" * 60, Color.CYAN)
        cprint("  开始批量预测", Color.CYAN, bold=True)
        cprint("─" * 60 + "\n", Color.CYAN)

        for idx, row in compounds_df.iterrows():
            smiles = str(row[smiles_col]).strip()
            if not smiles or smiles == "nan":
                logger.warning(f"第 {idx + 1} 行 SMILES 为空，跳过")
                continue

            compound_name = (
                str(row[name_col]).strip()
                if name_col
                and name_col in row.index
                and str(row[name_col]).strip() != "nan"
                else f"Compound_{idx + 1}"
            )

            # 检查批次冷却
            self._batch_cooldown_check()

            elapsed = int(time.time() - start_time)
            print_progress(idx, total, compound_name, "预测中...", elapsed)

            results = self.predict_single(smiles, compound_name)

            if results is not None:
                all_results[compound_name] = results
                print_progress(idx + 1, total, compound_name, f"✓ {len(results)}靶点", int(time.time() - start_time))
            else:
                all_results[compound_name] = []
                print_progress(idx + 1, total, compound_name, "✗ 失败", int(time.time() - start_time))

            # 等待间隔
            if idx < total - 1:
                delay = self.cfg.request_interval + random.uniform(*self.cfg.random_delay_range)
                # 等待时显示倒计时
                for t in range(int(delay), 0, -1):
                    print_progress(idx + 1, total, compound_name, f"等待{t}s...", int(time.time() - start_time))
                    time.sleep(1)

        # 最终换行
        elapsed = int(time.time() - start_time)
        print_progress(total, total, "完成!", "✓", elapsed)
        print()
        print()

        return all_results, elapsed

    def close(self):
        self._close_browser()


# ============================================================
# 输出函数
# ============================================================


def save_individual_results(all_results, output_dir):
    """保存每个化合物的单独 CSV（第1行 Compound，第2行 SMILES，第3行起为数据）"""
    os.makedirs(output_dir, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d")
    ind_dir = os.path.join(output_dir, f"individual_{ts}_results")
    os.makedirs(ind_dir, exist_ok=True)

    saved_count = 0
    for compound_name, results in all_results.items():
        if not results:
            continue
        df = pd.DataFrame(results)
        safe_name = re.sub(r'[\\/:*?"<>|]', '_', compound_name)
        filepath = os.path.join(ind_dir, f"{safe_name}_swisstarget.csv")
        try:
            smiles = results[0].get("SMILES", "") if results else ""
            with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                # 第1行：Compound 元数据
                writer.writerow(["Compound", compound_name])
                # 第2行：SMILES 元数据
                writer.writerow(["SMILES", smiles])
                # 第3行：列标题；第4行起：数据
                writer.writerow(list(df.columns))
                for _, row in df.iterrows():
                    writer.writerow(row.tolist())
            saved_count += 1
        except Exception as e:
            logger.error(f"{compound_name} CSV 导出失败: {e}")

    logger.info(f"保存了 {saved_count} 个单独文件")
    return saved_count


def save_merged_results(all_results, output_dir):
    """保存合并结果为 XLSX 格式（第1行 Date，第2行 Database，第3行起为数据）"""
    os.makedirs(output_dir, exist_ok=True)

    all_rows = []
    for compound_name, results in all_results.items():
        if results:
            all_rows.extend(results)

    if not all_rows:
        logger.warning("没有预测结果可合并")
        return None

    df_merged = pd.DataFrame(all_rows)
    df_merged.insert(0, "No.", range(1, len(df_merged) + 1))

    ts = datetime.now().strftime("%Y%m%d")
    filepath = os.path.join(output_dir, f"swisstarget_{ts}_merged_results.xlsx")
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Date", datetime.now().strftime("%Y-%m-%d")])
        ws.append(["Database", "SwissTargetPrediction"])
        # 写入列标题
        col_labels = list(df_merged.columns)
        ws.append(col_labels)
        # 写入数据行
        for _, row in df_merged.iterrows():
            ws.append(row.tolist())
        wb.save(filepath)
        logger.info(f"合并结果: {filepath} ({len(df_merged)} 条)")
    except Exception as e:
        logger.error(f"合并 Excel 导出失败: {e}")
        return None

    return filepath


def save_summary_report(all_results, output_dir):
    """保存汇总报告"""
    os.makedirs(output_dir, exist_ok=True)

    summary_data = []
    for compound_name, results in all_results.items():
        if results:
            df = pd.DataFrame(results)
            summary_data.append({
                "化合物名称": compound_name,
                "SMILES": results[0]["SMILES"][:100] if results else "",
                "预测靶点数": len(results),
                "最高概率": df["Probability"].max(),
                "平均概率": round(df["Probability"].mean(), 4),
                "靶点分类（唯一）": ", ".join(sorted(df["Target_Class"].unique())[:5]),
            })
        else:
            summary_data.append({
                "化合物名称": compound_name,
                "SMILES": "",
                "预测靶点数": 0,
                "最高概率": 0,
                "平均概率": 0,
                "靶点分类（唯一）": "无结果",
            })

    df_summary = pd.DataFrame(summary_data)
    ts = datetime.now().strftime("%Y%m%d")
    filepath = os.path.join(output_dir, f"swisstarget_{ts}_summary.xlsx")
    df_summary.to_excel(filepath, index=False, engine="openpyxl")
    logger.info(f"汇总报告: {filepath}")


def save_invalid_smiles(invalid_list, output_dir):
    """保存无效 SMILES 报告"""
    if not invalid_list:
        return None
    os.makedirs(output_dir, exist_ok=True)
    df_invalid = pd.DataFrame(invalid_list)
    ts = datetime.now().strftime("%Y%m%d")
    filepath = os.path.join(output_dir, f"swisstarget_{ts}_invalid_smiles.xlsx")
    df_invalid.to_excel(filepath, index=False, engine="openpyxl")
    logger.info(f"无效 SMILES 报告: {filepath}")
    return filepath


# ============================================================
# 交互式主程序
# ============================================================


def run_interactive(cfg, raw_df):
    """运行交互式预测流程"""
    global logger

    # 初始化日志
    logger = setup_logging(cfg.output_dir)

    # 提取 SMILES 列并过滤空值
    print_section("正在处理数据...")
    df = raw_df.copy()
    df = df.dropna(subset=[cfg.smiles_col])
    df = df[df[cfg.smiles_col].astype(str).str.strip() != ""]
    print_ok(f"有效数据行: {len(df)}")

    # SMILES 合法性预检
    print_section("SMILES 合法性预检...")
    validator_name = "RDKit（精确验证）" if HAS_RDKIT else "正则表达式（基础验证）"
    print_info(f"验证引擎: {validator_name}")
    valid_df, invalid_list = validate_smiles_batch(df, cfg.smiles_col)

    if invalid_list:
        print_warn(f"发现 {len(invalid_list)} 条无效 SMILES：")
        for item in invalid_list[:5]:  # 只显示前5条
            print_err(f"  行 {item['行号']}: {item['SMILES'][:40]}... → {item['原因']}")
        if len(invalid_list) > 5:
            print_info(f"  ... 还有 {len(invalid_list) - 5} 条，详见报告文件")

        # 保存无效 SMILES 报告
        invalid_path = save_invalid_smiles(invalid_list, cfg.output_dir)
        if invalid_path:
            print_ok(f"无效 SMILES 报告已保存: {os.path.basename(invalid_path)}")

    if len(valid_df) == 0:
        print_err("没有有效的 SMILES，程序退出")
        return {}

    print_ok(f"有效 SMILES: {len(valid_df)} 条")
    print()

    # 预估时间
    avg_time = cfg.request_interval + sum(cfg.random_delay_range) / 2 + 25  # 25秒预测处理
    est_total = len(valid_df) * avg_time
    print_info(f"预计耗时: {est_total / 60:.1f} 分钟（{len(valid_df)} 个化合物 × ~{avg_time:.0f}秒/个）\n")

    # 创建预测器
    predictor = SwissTargetPredictor(cfg)
    all_results = {}
    success = False

    try:
        all_results, elapsed = predictor.batch_predict(valid_df, cfg.smiles_col, cfg.name_col)
        success = True

    except KeyboardInterrupt:
        print()
        print()
        cprint("╔" + "═" * 58 + "╗", Color.YELLOW, bold=True)
        cprint("║" + "  ⚠ 程序已被用户中断".center(50) + "║", Color.YELLOW, bold=True)
        cprint("╚" + "═" * 58 + "╝", Color.YELLOW, bold=True)
        print()

        # 保存已完成的部分
        if all_results:
            print_section("正在保存已完成的预测结果...")
            do_save_results(all_results, cfg.output_dir)
            total_targets = sum(len(r) for r in all_results.values())
            success_count = sum(1 for r in all_results.values() if r)
            print_ok(f"已保存 {success_count}/{len(valid_df)} 个化合物的结果（{total_targets} 个靶点）")
        else:
            print_warn("尚无完成的预测结果")
        print()

    except Exception as e:
        print()
        print_err(f"程序运行出错: {e}")
        traceback.print_exc()
        print()

        if all_results:
            print_warn("正在保存已完成的预测结果...")
            do_save_results(all_results, cfg.output_dir)

    finally:
        predictor.close()

    # 显示最终统计
    if success and all_results:
        print_section("运行完成")
        print()

        total_targets = sum(len(r) for r in all_results.values())
        success_count = sum(1 for r in all_results.values() if r)
        fail_count = sum(1 for r in all_results.values() if not r)

        cprint("  ──────────── 运行统计 ────────────", Color.BOLD)
        print(f"  总化合物数:      {len(valid_df)}")
        print(f"  成功预测数:      {success_count}")
        if fail_count > 0:
            print_warn(f"  预测失败数:      {fail_count}")
        print(f"  总靶点数:        {total_targets}")
        if 'elapsed' in dir() and elapsed:
            print(f"  总耗时:          {elapsed / 60:.1f} 分钟")
        else:
            now = time.time()
            if hasattr(time, 'start_time'):
                pass
        print(f"  输出目录:        {os.path.abspath(cfg.output_dir)}")
        cprint("  ──────────────────────────────────", Color.BOLD)
        print()

        # 保存结果
        print_section("正在保存结果...")
        do_save_results(all_results, cfg.output_dir)
        print()

    return all_results


def do_save_results(all_results, output_dir):
    """执行保存操作"""
    ts = datetime.now().strftime("%Y%m%d")
    saved = save_individual_results(all_results, output_dir)
    print_ok(f"单独化合物文件: {saved} 个 → individual_{ts}_results/")
    merged_path = save_merged_results(all_results, output_dir)
    if merged_path:
        print_ok(f"合并结果文件: {os.path.basename(merged_path)}")
    save_summary_report(all_results, output_dir)
    print_ok("汇总报告已保存")


def post_run_menu(cfg):
    """运行结束后的交互菜单"""
    while True:
        print()
        print_section("下一步操作")
        print("  [1] 重新运行（使用相同输入文件）")
        print("  [2] 重新运行（修改配置）")
        print("  [3] 在文件管理器中打开结果目录")
        print("  [4] 退出")
        print()
        choice = input("  请选择 [4]: ").strip()

        if choice == "1":
            return "rerun_same"
        elif choice == "2":
            return "rerun_new"
        elif choice == "3":
            try:
                abs_path = os.path.abspath(cfg.output_dir)
                os.startfile(abs_path)
                print_ok(f"已打开文件夹: {abs_path}")
            except Exception as e:
                print_err(f"无法打开文件夹: {e}")
        else:
            return "exit"


# ============================================================
# 入口
# ============================================================


def main():
    """主入口"""
    while True:
        # 交互式配置
        cfg, raw_df = interactive_setup()

        # 运行预测
        run_interactive(cfg, raw_df)

        # 运行后菜单
        action = post_run_menu(cfg)
        if action == "exit":
            break
        elif action == "rerun_new":
            continue
        elif action == "rerun_same":
            continue

    print()
    cprint("感谢使用 SwissTargetPrediction 自动化预测工具！", Color.GREEN, bold=True)
    print()


if __name__ == "__main__":
    main()
