#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TargetNet 交互式化合物靶点预测脚本  v3.0
========================================
基于终端交互式操作，自动提交 SMILES 或 .smi 文件到 TargetNet
并解析返回的靶点预测结果，支持 CSV/Excel 导出。

【v3.0 更新】
  - 新增直接读取 Excel 文件，用户指定 SMILES 列和名称列后自动转换
  - 输出结果默认存入 targetnet_results 文件夹（可自定义路径）
  - 批量预测完成后自动生成合并结果表格
  - 保留参数预设、单分子/批量模式

依赖: pip install requests beautifulsoup4 pandas openpyxl colorama
"""

import os
import re
import sys
import csv
import json
import textwrap
from pathlib import Path
from datetime import datetime

# ─── tkinter (GUI) ──────────────────────────────────────
try:
    import tkinter as tk
    from tkinter import ttk, messagebox
    HAS_TK = True
except ImportError:
    HAS_TK = False

# ─── 可选依赖 ───────────────────────────────────────────
try:
    import requests
except ImportError:
    print("[错误] 缺少 requests 库，请运行: pip install requests")
    sys.exit(1)

try:
    from bs4 import BeautifulSoup
    HAS_BS4 = True
except ImportError:
    HAS_BS4 = False

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    from colorama import init, Fore, Style
    init(autoreset=True)
    HAS_COLOR = True
except ImportError:
    HAS_COLOR = False

# ─── 路径与配置文件 ─────────────────────────────────────
SCRIPT_DIR = Path(__file__).resolve().parent
CONFIG_FILE = SCRIPT_DIR / "targetnet_presets.json"

# ─── 常量 ───────────────────────────────────────────────
BASE_URL = "http://targetnet.scbdd.com"
SINGLE_ENDPOINT = f"{BASE_URL}/calcnet/calc_text/"
BATCH_ENDPOINT = f"{BASE_URL}/calcnet/calc_list/"
TIMEOUT_SINGLE = 120
TIMEOUT_BATCH  = 600
BATCH_LIMIT = 5000

FINGERPRINT_MAP = {
    "ecfp4":    "BaseInfo_ecfp4",
    "ecfp6":    "BaseInfo_ecfp6",
    "ecfp2":    "BaseInfo_ecfp2",
    "fp2":      "BaseInfo_fp2",
    "maccs":    "BaseInfo_maccs",
    "estate":   "BaseInfo_estate",
    "daylight": "BaseInfo_daylight",
}
FINGERPRINT_LABELS = {
    "ecfp4":    "ECFP4 fingerprints (默认推荐)",
    "ecfp6":    "ECFP6 fingerprints",
    "ecfp2":    "ECFP2 fingerprints",
    "fp2":      "FP2 fingerprints",
    "maccs":    "MACCS fingerprints",
    "estate":   "Estate fingerprints",
    "daylight": "Daylight fingerprints",
}

METRIC_MAP = {
    "auc": "AUC_score",
    "acc": "ACC",
    "mcc": "MCC",
    "f1":  "F1_score",
}
METRIC_LABELS = {
    "auc": "AUC >= (推荐)",
    "acc": "Accuracy >=",
    "mcc": "MCC >=",
    "f1":  "F-score >=",
}

# ─── 默认预设 ───────────────────────────────────────────
DEFAULT_PRESETS = {
    "fingerprint": "ecfp4",
    "metric": "auc",
    "threshold": 0.7,
    "auto_export": "csv",
    "max_show": 30,
}


# ═══════════════════════════════════════════════════════════
#  终端样式
# ═══════════════════════════════════════════════════════════
def c(text, color="white", bold=False):
    if not HAS_COLOR:
        return text
    codes = {"r": Fore.RED, "g": Fore.GREEN, "y": Fore.YELLOW,
             "b": Fore.BLUE, "c": Fore.CYAN, "m": Fore.MAGENTA, "w": Fore.WHITE}
    prefix = codes.get(color, "")
    if bold: prefix += Style.BRIGHT
    return f"{prefix}{text}{Style.RESET_ALL}"

def print_banner():
    print(c(r"""
╔══════════════════════════════════════════════════════╗
║           TargetNet 化合物靶点预测工具 v3.0            ║
║     http://targetnet.scbdd.com/calcnet/index/         ║
║     基于 QSAR 模型的交互式靶点预测 · 支持Excel输入     ║
╚══════════════════════════════════════════════════════╝""", "c", bold=True))

def print_info(msg):   print(f"  {c('[i]', 'b')} {msg}")
def print_ok(msg):     print(f"  {c('[√]', 'g')} {msg}")
def print_warn(msg):   print(f"  {c('[!]', 'y')} {msg}")
def print_error(msg):  print(f"  {c('[×]', 'r')} {msg}")
def print_section(t):  print(f"\n  {c('─'*52, 'c')}\n  {c(t, 'y', bold=True)}\n  {c('─'*52, 'c')}")


# ═══════════════════════════════════════════════════════════
#  配置文件管理
# ═══════════════════════════════════════════════════════════
def load_config():
    """加载配置，不存在则创建默认"""
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            # 合并默认值，防止配置版本不兼容
            merged = dict(DEFAULT_PRESETS)
            merged.update(cfg)
            return merged
        except (json.JSONDecodeError, KeyError):
            print_warn("配置文件损坏，已重置为默认值")
    save_config(DEFAULT_PRESETS)
    return dict(DEFAULT_PRESETS)


def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def show_current_preset():
    cfg = load_config()
    fp = cfg.get("fingerprint", "ecfp4")
    mt = cfg.get("metric", "auc")
    th = cfg.get("threshold", 0.7)
    ex = cfg.get("auto_export", "csv")
    print_section("当前预设参数")
    print(f"  分子指纹 : {c(FINGERPRINT_LABELS.get(fp, fp), 'g')}")
    print(f"  筛选指标 : {c(METRIC_LABELS.get(mt, mt), 'g')}")
    print(f"  阈值     : {c(str(th), 'g')}")
    print(f"  自动导出 : {c(ex, 'g')}")


def interactive_settings():
    """预设参数设置"""
    cfg = load_config()
    fp_keys = list(FINGERPRINT_MAP.keys())
    mt_keys = list(METRIC_MAP.keys())

    while True:
        show_current_preset()
        print(f"\n  {c('[1]', 'y')} 修改分子指纹")
        print(f"  {c('[2]', 'y')} 修改筛选指标")
        print(f"  {c('[3]', 'y')} 修改筛选阈值")
        print(f"  {c('[4]', 'y')} 修改自动导出格式")
        print(f"  {c('[5]', 'y')} 恢复默认预设")
        print(f"  {c('[0]', 'r')} 返回主菜单")

        ch = input(f"\n  请选择 [0-5]: ").strip()
        if ch == "1":
            print()
            for i, k in enumerate(fp_keys, 1):
                mk = " *" if k == cfg["fingerprint"] else "  "
                print(f"  {mk} [{c(str(i), 'y')}] {FINGERPRINT_LABELS[k]}")
            sel = input(f"\n  选择指纹 [{fp_keys.index(cfg['fingerprint'])+1}]: ").strip()
            try:
                idx = int(sel) - 1
                if 0 <= idx < len(fp_keys):
                    cfg["fingerprint"] = fp_keys[idx]
                    save_config(cfg)
                    print_ok("已更新")
            except ValueError:
                print_error("无效选项")
        elif ch == "2":
            print()
            for i, k in enumerate(mt_keys, 1):
                mk = " *" if k == cfg["metric"] else "  "
                print(f"  {mk} [{c(str(i), 'y')}] {METRIC_LABELS[k]}")
            sel = input(f"\n  选择指标 [{mt_keys.index(cfg['metric'])+1}]: ").strip()
            try:
                idx = int(sel) - 1
                if 0 <= idx < len(mt_keys):
                    cfg["metric"] = mt_keys[idx]
                    save_config(cfg)
                    print_ok("已更新")
            except ValueError:
                print_error("无效选项")
        elif ch == "3":
            val = input(f"\n  输入阈值 (0.1~1.0) [{cfg['threshold']}]: ").strip()
            if not val:
                continue
            try:
                f = float(val)
                if 0.1 <= f <= 1.0:
                    cfg["threshold"] = f
                    save_config(cfg)
                    print_ok("已更新")
                else:
                    print_error("阈值需在 0.1~1.0 之间")
            except ValueError:
                print_error("请输入有效数字")
        elif ch == "4":
            print()
            print("  [1] csv   [2] xlsx   [3] txt   [4] 不自动导出")
            cur = {"csv": "1", "xlsx": "2", "txt": "3", "none": "4"}.get(cfg.get("auto_export", "csv"), "1")
            sel = input(f"  选择格式 [{cur}]: ").strip() or cur
            m = {"1": "csv", "2": "xlsx", "3": "txt", "4": "none"}
            if sel in m:
                cfg["auto_export"] = m[sel]
                save_config(cfg)
                print_ok("已更新")
        elif ch == "5":
            save_config(DEFAULT_PRESETS)
            cfg = dict(DEFAULT_PRESETS)
            print_ok("已恢复默认预设")
        elif ch == "0":
            break
        else:
            print_error("无效选项")


# ═══════════════════════════════════════════════════════════
#  Excel 读取与转换
# ═══════════════════════════════════════════════════════════
def read_excel_columns(filepath):
    """读取 Excel 文件的列名列表"""
    if not HAS_PANDAS:
        print_error("需要 pandas 和 openpyxl 来读取 Excel 文件")
        print_info("安装: pip install pandas openpyxl")
        return None
    try:
        df = pd.read_excel(filepath, nrows=0, engine="openpyxl")
        return list(df.columns)
    except Exception as e:
        print_error(f"读取 Excel 失败: {e}")
        return None


def excel_to_smi(excel_path, smiles_col, name_col=None, output_smi=None):
    """
    将 Excel 中指定列转换为 .smi 文件。
    返回: (smi_path, compounds_list)  compounds_list = [(smiles, name), ...]
    """
    if not HAS_PANDAS:
        print_error("需要 pandas 和 openpyxl")
        return None, None

    try:
        df = pd.read_excel(excel_path, engine="openpyxl")
    except Exception as e:
        print_error(f"读取 Excel 失败: {e}")
        return None, None

    # 检查列是否存在
    if smiles_col not in df.columns:
        print_error(f"SMILES 列 '{smiles_col}' 不存在，可用列: {list(df.columns)}")
        return None, None
    if name_col and name_col not in df.columns:
        print_warn(f"名称列 '{name_col}' 不存在，将自动编号")
        name_col = None

    # 过滤空值
    df = df.dropna(subset=[smiles_col])
    if df.empty:
        print_error(f"SMILES 列 '{smiles_col}' 无有效数据")
        return None, None

    compounds = []
    for _, row in df.iterrows():
        smi = str(row[smiles_col]).strip()
        name = str(row[name_col]).strip() if name_col and pd.notna(row.get(name_col)) else ""
        if smi and smi.lower() != "nan":
            compounds.append((smi, name))

    if not compounds:
        print_error("未提取到有效 SMILES 数据")
        return None, None

    # 写 .smi 文件
    if output_smi is None:
        output_smi = Path(excel_path).with_suffix(".smi")
    with open(output_smi, "w", encoding="utf-8") as f:
        for smi, name in compounds:
            if name:
                f.write(f"{smi} {name}\n")
            else:
                f.write(f"{smi}\n")

    print_ok(f"从 Excel 提取 {c(str(len(compounds)), 'g')} 个化合物 → {Path(output_smi).name}")
    return str(output_smi), compounds


def interactive_excel_select(excel_path):
    """
    通过 GUI 弹窗选择 Excel 中的 SMILES 列和名称列。
    若 tkinter 不可用则降级为终端交互。
    返回: (smiles_col_name, name_col_name_or_None)
    """
    columns = read_excel_columns(excel_path)
    if columns is None:
        return None, None

    if HAS_TK:
        return _gui_excel_select(excel_path, columns)
    else:
        print_warn("tkinter 不可用，降级为终端选择模式")
        return _terminal_excel_select(excel_path, columns)


def _gui_excel_select(excel_path, columns):
    """GUI 弹窗选择 SMILES 列和名称列"""
    result = {"smiles_col": None, "name_col": None, "confirmed": False}

    root = tk.Tk()
    root.title("TargetNet - Excel 列选择")
    root.geometry("480x360")
    root.resizable(False, False)
    # 居中显示
    root.update_idletasks()
    x = (root.winfo_screenwidth() - 480) // 2
    y = (root.winfo_screenheight() - 360) // 2
    root.geometry(f"+{x}+{y}")
    root.attributes("-topmost", True)
    root.focus_force()

    # ── 标题 ──
    tk.Label(root, text="Excel 列选择",
             font=("Microsoft YaHei UI", 14, "bold")).pack(pady=(18, 4))
    # 作者/站点信息
    info_frame = tk.Frame(root)
    info_frame.pack(pady=(0, 2))
    tk.Label(info_frame, text="author: shenjianlin", font=("Microsoft YaHei UI", 8), fg="#888").pack(side="left", padx=5)
    tk.Label(info_frame, text="site: git@github.com:sjl-openmywork/Network-toxicology-script.git", font=("Microsoft YaHei UI", 8), fg="#888").pack(side="left", padx=5)
    tk.Label(root, text=f"文件: {Path(excel_path).name}",
             font=("Microsoft YaHei UI", 9), fg="#555").pack()
    tk.Label(root, text=f"共 {len(columns)} 列",
             font=("Microsoft YaHei UI", 9), fg="#888").pack(pady=(0, 10))

    # ── SMILES 列选择 ──
    frame1 = tk.Frame(root)
    frame1.pack(fill="x", padx=30, pady=6)
    tk.Label(frame1, text="SMILES 列:", font=("Microsoft YaHei UI", 10)).pack(side="left")
    smiles_var = tk.StringVar()
    smiles_combo = ttk.Combobox(frame1, textvariable=smiles_var,
                                values=columns, state="readonly", width=28)
    smiles_combo.pack(side="left", padx=(10, 0))
    # 自动猜测 SMILES 列
    guessed = _guess_smiles_column(columns)
    smiles_combo.current(guessed)

    # ── 名称列选择 ──
    frame2 = tk.Frame(root)
    frame2.pack(fill="x", padx=30, pady=6)
    tk.Label(frame2, text="名称列:  ", font=("Microsoft YaHei UI", 10)).pack(side="left")
    name_var = tk.StringVar(value="(不使用)")
    name_options = ["(不使用)"] + list(columns)
    name_combo = ttk.Combobox(frame2, textvariable=name_var,
                              values=name_options, state="readonly", width=28)
    name_combo.pack(side="left", padx=(10, 0))
    name_combo.current(0)

    # ── 数据预览 ──
    tk.Label(root, text="数据预览 (前3行):", font=("Microsoft YaHei UI", 9), fg="#666").pack(pady=(12, 2))
    preview_frame = tk.Frame(root)
    preview_frame.pack(fill="both", padx=30, expand=False)
    try:
        df_preview = pd.read_excel(excel_path, nrows=3, engine="openpyxl")
        cols_for_preview = list(df_preview.columns)
        tree = ttk.Treeview(preview_frame, columns=cols_for_preview, show="headings", height=3)
        for col_name in cols_for_preview:
            tree.heading(col_name, text=str(col_name))
            tree.column(col_name, width=max(80, 400 // len(cols_for_preview)))
        for _, row in df_preview.iterrows():
            tree.insert("", "end", values=[str(v)[:40] for v in row])
        tree.pack(fill="x")
    except Exception:
        tk.Label(preview_frame, text="(预览失败)", fg="#999").pack()

    # ── 确认按钮 ──
    def on_confirm():
        result["smiles_col"] = smiles_var.get()
        nv = name_var.get()
        result["name_col"] = None if nv == "(不使用)" else nv
        result["confirmed"] = True
        root.destroy()

    def on_cancel():
        root.destroy()

    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=14)
    tk.Button(btn_frame, text="  确认  ", command=on_confirm,
              font=("Microsoft YaHei UI", 10), bg="#4CAF50", fg="white",
              width=10, height=1).pack(side="left", padx=8)
    tk.Button(btn_frame, text="  取消  ", command=on_cancel,
              font=("Microsoft YaHei UI", 10), width=10, height=1).pack(side="left", padx=8)

    root.protocol("WM_DELETE_WINDOW", on_cancel)
    root.mainloop()

    if result["confirmed"] and result["smiles_col"]:
        sc = result["smiles_col"]
        nc = result["name_col"]
        label = f"SMILES 列: {c(sc, 'g')}" + (f" | 名称列: {c(nc, 'g')}" if nc else "")
        print_ok(label)
        return sc, nc
    else:
        print_warn("已取消列选择")
        return None, None


def _guess_smiles_column(columns):
    """自动猜测哪一列是 SMILES 列，返回索引"""
    keywords = ["smiles", "smi", "structure", "mol"]
    for i, col in enumerate(columns):
        if col.lower().strip() in keywords:
            return i
    # 没找到关键词，默认第1列
    return 0


def _terminal_excel_select(excel_path, columns):
    """终端降级模式选择列"""
    print_section("Excel 列选择 (终端模式)")
    print_info(f"文件: {Path(excel_path).name}")
    print_info(f"检测到 {c(str(len(columns)), 'g')} 列:")
    print()
    for i, col in enumerate(columns, 1):
        print(f"  [{c(str(i), 'y')}] {col}")

    print()
    sel = input(f"  请选择 SMILES 列序号 [1]: ").strip() or "1"
    try:
        smi_idx = int(sel) - 1
        if not (0 <= smi_idx < len(columns)):
            raise ValueError
    except ValueError:
        print_error("无效选择，使用第1列")
        smi_idx = 0
    smiles_col = columns[smi_idx]

    print()
    sel2 = input(f"  请选择名称列序号 (直接回车跳过): ").strip()
    name_col = None
    if sel2:
        try:
            name_idx = int(sel2) - 1
            if 0 <= name_idx < len(columns):
                name_col = columns[name_idx]
            else:
                print_warn("无效选择，跳过名称列")
        except ValueError:
            print_warn("无效输入，跳过名称列")

    print_ok(f"SMILES 列: {c(smiles_col, 'g')}" + (f" | 名称列: {c(name_col, 'g')}" if name_col else ""))
    return smiles_col, name_col


# ═══════════════════════════════════════════════════════════
#  输出目录管理
# ═══════════════════════════════════════════════════════════
def get_output_dir(user_specified=None):
    """
    获取输出目录。默认 SCRIPT_DIR/targetnet_results，可用户指定。
    自动创建目录。
    """
    if user_specified:
        out = Path(user_specified)
    else:
        out = SCRIPT_DIR / "targetnet_results"
    out.mkdir(parents=True, exist_ok=True)
    return out


def merge_results_to_table(all_records, output_dir):
    """
    将所有化合物的预测结果合并为一个 XLSX 文件。
    输出格式：第1行 Date，第2行 Database，第3行起为数据。
    """
    if not all_records:
        print_warn("无结果可合并")
        return []

    if not HAS_PANDAS or not HAS_OPENPYXL:
        print_warn("缺少 pandas/openpyxl，无法生成合并 Excel")
        return []

    fieldnames = ["compound_name", "smiles", "uniprot_id", "protein", "probability"]
    col_labels = ["Compound_Name", "SMILES", "UniProt_ID", "Protein", "Probability"]
    ts = datetime.now().strftime("%Y%m%d")

    xlsx_path = output_dir / f"targetnet_{ts}_merged_results.xlsx"
    try:
        df = pd.DataFrame(all_records)
        df = df[fieldnames]
        df.columns = col_labels

        # 使用 openpyxl 写入带元数据头部的 XLSX
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Date", datetime.now().strftime("%Y-%m-%d")])
        ws.append(["Database", "TargetNet"])
        # 写入列标题
        ws.append(col_labels)
        # 写入数据行
        for _, row in df.iterrows():
            ws.append(row.tolist())
        wb.save(xlsx_path)
        return [str(xlsx_path)]
    except Exception as e:
        print_error(f"合并 Excel 导出失败: {e}")

    return []


def export_individual_results(all_records, output_dir):
    """
    按化合物分组，为每个化合物单独导出 CSV 结果文件。
    CSV 格式：第1行 Compound，第2行 SMILES，第3行起为预测数据。
    文件存入 individual_{YYYYMMDD}_results/ 子目录。
    返回: 保存的文件路径列表
    """
    if not all_records:
        return []

    ts = datetime.now().strftime("%Y%m%d")

    # 按化合物名分组
    compound_groups = {}
    for rec in all_records:
        name = rec.get("compound_name", "") or "Unknown"
        if name not in compound_groups:
            compound_groups[name] = []
        compound_groups[name].append(rec)

    # 创建带日期的子文件夹
    ind_dir = output_dir / f"individual_{ts}_results"
    ind_dir.mkdir(parents=True, exist_ok=True)

    saved = []
    for name, records in compound_groups.items():
        # 清理文件名
        safe_name = re.sub(r'[<>:"/\\|?*]', '_', name)[:50]
        fp = ind_dir / f"{safe_name}_targetnet.csv"

        try:
            with open(fp, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                # 第1行：Compound 元数据
                smiles = records[0].get("smiles", "") if records else ""
                writer.writerow(["Compound", name])
                # 第2行：SMILES 元数据
                writer.writerow(["SMILES", smiles])
                # 第3行：列标题
                writer.writerow(["UniProt_ID", "Protein", "Probability"])
                # 第4行起：数据
                for rec in records:
                    writer.writerow([
                        rec.get("uniprot_id", ""),
                        rec.get("protein", ""),
                        rec.get("probability", ""),
                    ])
            saved.append(str(fp))
        except Exception as e:
            print_error(f"{name} CSV 导出失败: {e}")

    return saved


# ═══════════════════════════════════════════════════════════
#  SMI 文件解析（支持含名称行）
# ═══════════════════════════════════════════════════════════
def parse_smi_file(filepath):
    """
    解析 .smi 文件，支持：
      SMILES
      SMILES NAME
      SMILES\tNAME
    返回: [(smiles, name), ...]
    """
    compounds = []
    with open(filepath, "r", encoding="utf-8", errors="replace") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            parts = line.split(None, 1)  # 按空白分割，最多2部分
            if len(parts) == 1:
                compounds.append((parts[0], ""))
            else:
                compounds.append((parts[0], parts[1]))
    return compounds


# ═══════════════════════════════════════════════════════════
#  提交 & 解析
# ═══════════════════════════════════════════════════════════
def submit_single(smiles, fingerprint_code, metric_code, threshold):
    """提交单个 SMILES"""
    payload = {
        "smile": smiles,
        "param_1": METRIC_MAP[metric_code],
        "param_2": str(threshold),
        "finger_type": FINGERPRINT_MAP[fingerprint_code],
    }
    print_info(f"提交: {smiles[:55]}{'...' if len(smiles)>55 else ''}")
    print_info(f"指纹: {fingerprint_code} | {metric_code}>={threshold} | 等待服务器...")

    try:
        resp = requests.post(SINGLE_ENDPOINT, data=payload, timeout=TIMEOUT_SINGLE)
        resp.raise_for_status()
        return resp.text
    except requests.Timeout:
        print_error(f"请求超时 (>{TIMEOUT_SINGLE}s)")
        return None
    except requests.RequestException as e:
        print_error(f"请求失败: {e}")
        return None


def submit_batch(filepath, fingerprint_code, metric_code, threshold):
    """提交批量 .smi 文件"""
    path = Path(filepath)
    if not path.exists():
        print_error(f"文件不存在: {filepath}")
        return None

    # 统计行数
    compounds = parse_smi_file(filepath)
    count = len(compounds)
    if count > BATCH_LIMIT:
        print_error(f"分子数 {count} 超过上限 {BATCH_LIMIT}，请拆分文件")
        return None

    print_info(f"提交文件: {path.name} ({count} 分子)")
    print_info(f"指纹: {fingerprint_code} | {metric_code}>={threshold} | 等待服务器...")

    try:
        with open(filepath, "rb") as f:
            resp = requests.post(
                BATCH_ENDPOINT,
                files={"tempfile": (path.name, f, "text/plain")},
                data={
                    "param_1_list": METRIC_MAP[metric_code],
                    "param_2_list": str(threshold),
                    "finger_type_list": FINGERPRINT_MAP[fingerprint_code],
                },
                timeout=TIMEOUT_BATCH,
            )
        resp.raise_for_status()
        return resp.text, compounds
    except requests.Timeout:
        print_error(f"请求超时 (>{TIMEOUT_BATCH}s)，请减少分子数")
        return None, None
    except requests.RequestException as e:
        print_error(f"请求失败: {e}")
        return None, None


# ─── 自动分批提交阈值 ───────────────────────────────────
AUTO_SPLIT_THRESHOLD = 15   # 超过此数量自动分批
AUTO_SPLIT_SIZE = 10        # 每批分子数


def submit_batch_auto(compounds, fingerprint_code, metric_code, threshold, output_dir=None):
    """
    自动分批提交化合物列表。
    化合物数 <= AUTO_SPLIT_THRESHOLD 时直接提交，否则分批。
    返回: 所有解析后的 records 列表
    """
    count = len(compounds)

    # 小批量直接提交
    if count <= AUTO_SPLIT_THRESHOLD:
        # 写临时 .smi
        tmp_smi = (output_dir or SCRIPT_DIR) / "_tmp_submit.smi"
        with open(tmp_smi, "w", encoding="utf-8") as f:
            for smi, name in compounds:
                f.write(f"{smi} {name}\n" if name else f"{smi}\n")
        html, _ = submit_batch(str(tmp_smi), fingerprint_code, metric_code, threshold)
        tmp_smi.unlink(missing_ok=True)
        if html:
            return parse_batch_html(html, compounds)
        return []

    # 大批量分批提交
    total_batches = (count + AUTO_SPLIT_SIZE - 1) // AUTO_SPLIT_SIZE
    print_info(f"化合物 {c(str(count), 'g')} 个，自动分为 {c(str(total_batches), 'g')} 批 (每批 {AUTO_SPLIT_SIZE} 个)")

    all_records = []
    fail_count = 0
    save_dir = output_dir or SCRIPT_DIR

    for batch_idx in range(total_batches):
        start = batch_idx * AUTO_SPLIT_SIZE
        end = min(start + AUTO_SPLIT_SIZE, count)
        batch_compounds = compounds[start:end]

        print(f"\n  {c(f'[批次 {batch_idx+1}/{total_batches}]', 'c')} 化合物 {start+1}-{end}")

        # 写临时 .smi
        batch_smi = save_dir / f"_batch_{batch_idx+1}.smi"
        with open(batch_smi, "w", encoding="utf-8") as f:
            for smi, name in batch_compounds:
                f.write(f"{smi} {name}\n" if name else f"{smi}\n")

        html, _ = submit_batch(str(batch_smi), fingerprint_code, metric_code, threshold)
        batch_smi.unlink(missing_ok=True)

        if html:
            records = parse_batch_html(html, batch_compounds)
            print_ok(f"解析到 {len(records)} 条结果")
            all_records.extend(records)
        else:
            fail_count += 1
            print_warn(f"批次 {batch_idx+1} 失败，已跳过")

    if fail_count > 0:
        print_warn(f"共 {fail_count} 个批次提交失败")

    return all_records


def parse_single_html(html, smiles, compound_name=""):
    """解析单分子结果 HTML → 带 SMILES+NAME 的记录"""
    records = _parse_table(html)
    # 给每条绑定化合物信息
    for r in records:
        r["smiles"] = smiles
        r["compound_name"] = compound_name
    return records


def parse_batch_html(html, compounds):
    """解析批量结果 HTML → 反透视每列分子概率 → 带 SMILES+NAME 的记录"""
    records = _parse_batch_table_unpivot(html, compounds)
    return records


def _parse_batch_table_unpivot(html, compounds):
    """
    批量结果表格: 表头列为 Details | Uniprot_ID | Protein | prob_mol0 | prob_mol1 | ...
    反透视: 每个靶点行 × 每个分子 → 一条记录
    """
    if not HAS_BS4:
        return _parse_batch_fallback(html, compounds)

    soup = BeautifulSoup(html, "html.parser")
    tables = soup.find_all("table")
    records = []

    for table in tables:
        rows = table.find_all("tr")
        if not rows:
            continue

        # ── 解析表头，确认概率列位置 ──
        thead = table.find("thead")
        prob_col_indices = []  # [(col_index, molecule_index), ...]
        if thead:
            ths = thead.find_all("th")
            for ci, th in enumerate(ths):
                text = th.get_text(strip=True)
                if text.isdigit():
                    mol_idx = int(text)
                    prob_col_indices.append((ci, mol_idx))
        if not prob_col_indices:
            # 无 thead 时，从第4列开始按序分配
            # 先读一行确定总列数
            data_rows = []
            for row in rows:
                cols = row.find_all("td")
                if len(cols) >= 4:
                    data_rows.append(cols)
                    break
            if data_rows:
                total_cols = len(data_rows[0])
                prob_start = 3  # 前3列: Details, Uniprot, Protein
                for pi in range(prob_start, total_cols):
                    prob_col_indices.append((pi, pi - prob_start))

        # ── 解析数据行 ──
        for row in rows:
            cols = row.find_all("td")
            if len(cols) < 4:
                continue

            texts = [col.get_text(strip=True) for col in cols]

            # 找 UniProt ID
            uniprot = None
            for t in texts:
                if re.match(r"^[A-Z][0-9][A-Z0-9]{4,}$", t):
                    uniprot = t
                    break
            if not uniprot:
                continue

            # 找蛋白名称
            protein = ""
            for t in texts:
                if (t != uniprot and len(t) > 5
                        and not re.match(r"^[\d.]+$", t)
                        and not t.lower() == "view"):
                    protein = t
                    break

            # 每个分子列的概率
            for col_i, mol_idx in prob_col_indices:
                prob = ""
                if col_i < len(texts):
                    prob = texts[col_i]
                # 为每个分子生成记录
                smiles = ""
                name = ""
                if mol_idx < len(compounds):
                    smiles = compounds[mol_idx][0]
                    name = compounds[mol_idx][1]
                    if not name:
                        name = f"Compound_{mol_idx + 1}"

                records.append({
                    "compound_name": name,
                    "smiles": smiles,
                    "uniprot_id": uniprot,
                    "protein": protein,
                    "probability": prob,
                })

    return records


def _parse_batch_fallback(html, compounds):
    """无 bs4 时的批量降级解析"""
    # 用正则找所有概率列
    # 匹配: 一个 UniProt 行可能有多个后续概率
    uniprot_matches = list(re.finditer(r">([A-Z][0-9][A-Z0-9]{4,})<", html))
    prob_matches = list(re.finditer(r"<td[^>]*>\s*([0-9]+\.[0-9]+)\s*</td>", html))

    n_mols = len(compounds)
    records = []

    # 简化: 假设每个 uniprot 后面紧跟 n_mols 个概率值
    prob_idx = 0
    for um in uniprot_matches:
        uniprot = um.group(1)
        for mi in range(n_mols):
            if prob_idx >= len(prob_matches):
                break
            prob = prob_matches[prob_idx].group(1)
            name = compounds[mi][1] if mi < len(compounds) else ""
            if not name:
                name = f"Compound_{mi + 1}"
            smiles = compounds[mi][0] if mi < len(compounds) else ""
            records.append({
                "compound_name": name,
                "smiles": smiles,
                "uniprot_id": uniprot,
                "protein": "",
                "probability": prob,
            })
            prob_idx += 1

    return records


def _parse_table(html):
    """通用 HTML 表格解析（单分子结果）"""
    records = []
    if HAS_BS4:
        soup = BeautifulSoup(html, "html.parser")
        for table in soup.find_all("table"):
            for row in table.find_all("tr"):
                cols = row.find_all("td")
                if len(cols) < 3:
                    continue
                texts = [c.get_text(strip=True) for c in cols]
                uniprot = None
                for t in texts:
                    if re.match(r"^[A-Z][0-9][A-Z0-9]{4,}$", t):
                        uniprot = t
                        break
                if not uniprot:
                    continue
                protein = ""
                for t in texts:
                    if (t != uniprot and len(t) > 5
                            and not re.match(r"^[\d.]+$", t)
                            and t.lower() != "view"):
                        protein = t
                        break
                prob = ""
                for t in texts:
                    if re.match(r"^\d+\.\d+$", t):
                        prob = t
                        break
                if not prob and len(texts) >= 4:
                    prob = texts[-1]
                if uniprot:
                    records.append({
                        "compound_name": "",
                        "smiles": "",
                        "uniprot_id": uniprot,
                        "protein": protein,
                        "probability": prob,
                    })
    else:
        # 降级正则
        uniprots = re.findall(r">([A-Z][0-9][A-Z0-9]{4,})<", html)
        probs = re.findall(r"<td[^>]*>\s*([0-9]+\.[0-9]+)\s*</td>", html)
        for i in range(max(len(uniprots), len(probs))):
            records.append({
                "compound_name": "",
                "smiles": "",
                "uniprot_id": uniprots[i] if i < len(uniprots) else "",
                "protein": "",
                "probability": probs[i] if i < len(probs) else "",
            })
    return records


# ═══════════════════════════════════════════════════════════
#  显示 & 导出
# ═══════════════════════════════════════════════════════════
def display_results(records, max_show=30):
    """在终端显示带化合物信息的结果"""
    if not records:
        print_warn("未提取到靶点预测结果")
        return

    total = len(records)
    print_ok(f"共提取 {c(str(total), 'g', bold=True)} 条靶点预测记录")

    # 统计独特化合物
    compounds = set()
    for r in records:
        compounds.add((r.get("compound_name", ""), r.get("smiles", "")))
    if len(compounds) > 1:
        print_info(f"涉及 {c(str(len(compounds)), 'g')} 个化合物")
    elif len(compounds) == 1:
        name, smi = list(compounds)[0]
        label = name if name else smi[:40]
        print_info(f"化合物: {c(label[:50], 'g')}")

    # 概率分布
    try:
        high = sum(1 for r in records if float(r.get("probability", 0) or 0) >= 0.5)
        mid  = sum(1 for r in records if 0.1 <= float(r.get("probability", 0) or 0) < 0.5)
        low  = sum(1 for r in records if float(r.get("probability", 0) or 0) < 0.1)
        print_info(f"概率分布: 高≥0.5 {c(str(high),'r')} | 中0.1~0.5 {c(str(mid),'y')} | 低<0.1 {c(str(low),'c')}")
    except (ValueError, KeyError):
        pass

    print()
    # 多化合物模式时显示窄表
    multi = len(compounds) > 1
    if multi:
        hdr = f"  {'#':<4} {'Compound':<14} {'UniProt':<12} {'Prob':<8} {'Protein':<30}"
        sep = f"  {'─'*3} {'─'*13} {'─'*11} {'─'*7} {'─'*29}"
    else:
        hdr = f"  {'#':<4} {'UniProt ID':<13} {'Probability':<11} {'Protein':<40}"
        sep = f"  {'─'*3} {'─'*12} {'─'*10} {'─'*39}"

    print(c(hdr, "c", bold=True))
    print(c(sep, "c"))

    display_count = min(total, max_show)
    for i, rec in enumerate(records[:display_count], 1):
        prob_str = rec.get("probability", "")
        try:
            pv = float(prob_str)
            pd = c(f"{pv:<7.3f}", "r") if pv >= 0.7 else (c(f"{pv:<7.3f}", "y") if pv >= 0.3 else f"{pv:<7.3f}")
        except (ValueError, TypeError):
            pd = f"{prob_str:<7}"

        uniprot = rec.get("uniprot_id", "")[:12]
        protein = rec.get("protein", "")[:38]

        if multi:
            cname = rec.get("compound_name", "")[:13]
            print(f"  {i:<4} {cname:<14} {uniprot:<12} {pd} {protein[:28]}")
        else:
            print(f"  {i:<4} {uniprot:<13} {pd}   {protein}")

    if total > max_show:
        print(f"  ... 省略 {total - max_show} 条（导出文件含完整数据）")
    print()


def auto_export(records, base_name, fmt):
    """根据预设格式自动导出"""
    if not records or fmt == "none":
        return []

    formats = {"csv": ["csv"], "xlsx": ["xlsx"], "txt": ["txt"]}
    exts = formats.get(fmt, ["csv"])
    return _do_export(records, base_name, exts)


def interactive_export(records, base_name):
    """交互式选择导出格式"""
    if not records:
        return

    cfg = load_config()
    default = cfg.get("auto_export", "csv")
    def_fmt_label = {"csv": "1", "xlsx": "2", "txt": "3", "none": "5"}.get(default, "1")

    print_section("导出结果")
    print(f"  [1] CSV   [2] Excel   [3] TXT   [4] 全部   [5] 不导出")
    choice = input(f"\n  选择 [{def_fmt_label}]: ").strip() or def_fmt_label

    fmt_map = {"1": ["csv"], "2": ["xlsx"], "3": ["txt"], "4": ["csv", "xlsx", "txt"], "5": []}
    exts = fmt_map.get(choice, ["csv"])

    if not exts:
        print_info("跳过导出")
        return

    saved = _do_export(records, base_name, exts)
    if saved:
        print_ok("已保存:")
        for f in saved:
            print(f"      {c(f, 'g')}")
    else:
        print_warn("无文件被保存")


def _do_export(records, base_name, formats):
    """实际导出。CSV 使用3行头格式（Compound / SMILES / 列标题）。"""
    saved = []
    # 从记录中提取化合物名称和 SMILES
    compound_name = records[0].get("compound_name", "") if records else ""
    smiles = records[0].get("smiles", "") if records else ""

    for fmt in formats:
        if fmt == "csv":
            fp = f"{base_name}.csv"
            try:
                with open(fp, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    # 第1行：Compound 元数据
                    writer.writerow(["Compound", compound_name])
                    # 第2行：SMILES 元数据
                    writer.writerow(["SMILES", smiles])
                    # 第3行：列标题
                    writer.writerow(["UniProt_ID", "Protein", "Probability"])
                    # 第4行起：数据
                    for rec in records:
                        writer.writerow([
                            rec.get("uniprot_id", ""),
                            rec.get("protein", ""),
                            rec.get("probability", ""),
                        ])
                saved.append(fp)
            except Exception as e:
                print_error(f"CSV 导出失败: {e}")

        elif fmt == "xlsx":
            if HAS_PANDAS and HAS_OPENPYXL:
                fp = f"{base_name}.xlsx"
                try:
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.append(["Date", datetime.now().strftime("%Y-%m-%d")])
                    ws.append(["Database", "TargetNet"])
                    ws.append(["Compound_Name", "SMILES", "UniProt_ID", "Protein", "Probability"])
                    for rec in records:
                        ws.append([
                            rec.get("compound_name", ""),
                            rec.get("smiles", ""),
                            rec.get("uniprot_id", ""),
                            rec.get("protein", ""),
                            rec.get("probability", ""),
                        ])
                    wb.save(fp)
                    saved.append(fp)
                except Exception as e:
                    print_error(f"Excel 导出失败: {e}")
            else:
                print_warn("缺少 pandas/openpyxl，跳过 Excel")

        elif fmt == "txt":
            fp = f"{base_name}.txt"
            try:
                with open(fp, "w", encoding="utf-8") as f:
                    f.write(f"# TargetNet 靶点预测结果\n")
                    f.write(f"# 导出时间: {datetime.now():%Y-%m-%d %H:%M:%S}\n")
                    f.write(f"# 共 {len(records)} 条记录\n#\n")
                    f.write(f"{'Compound':<16}\t{'SMILES':<45}\t{'UniProt_ID':<13}\t{'Prob':<8}\t{'Protein'}\n")
                    f.write(f"{'─'*15}\t{'─'*44}\t{'─'*12}\t{'─'*7}\t{'─'*40}\n")
                    for r in records:
                        f.write(f"{r.get('compound_name',''):<16}\t"
                                f"{r.get('smiles','')[:44]:<45}\t"
                                f"{r.get('uniprot_id',''):<13}\t"
                                f"{r.get('probability',''):<8}\t"
                                f"{r.get('protein','')}\n")
                saved.append(fp)
            except Exception as e:
                print_error(f"TXT 导出失败: {e}")

    return saved


# ═══════════════════════════════════════════════════════════
#  交互式流程（简洁版 - 使用预设减少操作）
# ═══════════════════════════════════════════════════════════
def quick_single():
    """快速单分子预测 — 使用预设参数，结果存入输出目录"""
    cfg = load_config()

    print_section("快速单分子预测 (使用预设参数)")
    show_current_preset()
    print()

    # 输入 SMILES + 名称
    print(c("  输入化合物信息:", "y"))
    print(c("  格式: SMILES [名称]   (名称可选，空格分隔)", "c"))
    print(c("  提示: 输入 'example' 使用示例", "c"))

    # 读取输入（只取第一行非空内容作为 SMILES）
    line = ""
    while not line:
        raw = input("  > ").strip()
        if not raw:
            continue
        if raw.lower() == "example":
            line = "C(C=CC1)=C(C=1C(=O)O)O  Salicylic_acid"
            break
        line = raw

    parts = line.split(None, 1)
    smiles = parts[0]
    name = parts[1] if len(parts) > 1 else ""

    # 确认
    print()
    if name:
        print_info(f"化合物: {c(name, 'g')}")
    print_info(f"SMILES: {c(smiles[:60], 'g')}")
    confirm = input(f"\n  直接回车开始预测，输入 'c' 取消: ").strip()
    if confirm.lower() == "c":
        return

    # 提交
    print_section("提交预测")
    html = submit_single(smiles, cfg["fingerprint"], cfg["metric"], cfg["threshold"])
    if not html:
        return

    # 解析
    print_section("解析结果")
    records = parse_single_html(html, smiles, name)
    display_results(records, cfg.get("max_show", 30))

    # 导出到输出目录
    out_input = input(f"  输出目录 (回车默认 targetnet_results): ").strip().strip('"').strip("'")
    output_dir = get_output_dir(out_input if out_input else None)

    ts = datetime.now().strftime("%Y%m%d")
    ind_dir = output_dir / f"individual_{ts}_results"
    ind_dir.mkdir(parents=True, exist_ok=True)
    label = re.sub(r"[^a-zA-Z0-9_]", "_", name)[:30] if name else "single"
    base = str(ind_dir / f"{label}_targetnet")

    auto_fmt = cfg.get("auto_export", "csv")
    saved = auto_export(records, base, auto_fmt)
    if saved:
        print_ok("已自动导出:")
        for f in saved:
            print(f"      {c(f, 'g')}")


def quick_batch():
    """快速批量预测 — 支持 Excel / .smi 文件输入，结果存入输出目录"""
    cfg = load_config()

    print_section("快速批量预测 (使用预设参数)")
    show_current_preset()
    print()
    print_info("支持格式: Excel (.xlsx/.xls) 或 .smi 文件")
    print_info("Excel 文件将自动引导选择 SMILES 列和名称列")
    print_info(f"示例 .smi: {BASE_URL}/static/media/calcnet/example.smi")

    # 输入文件
    filepath = input("\n  请输入文件路径: ").strip().strip('"').strip("'")
    if not filepath:
        return
    path = Path(filepath)
    if not path.exists():
        print_error(f"文件不存在: {filepath}")
        return

    # ── 判断文件类型 ──
    smi_path = None
    compounds = None

    if path.suffix.lower() in (".xlsx", ".xls"):
        # Excel 文件 → 交互选择列 → 转 .smi
        smiles_col, name_col = interactive_excel_select(filepath)
        if not smiles_col:
            return
        temp_smi = path.with_suffix(".smi")
        smi_path, compounds = excel_to_smi(filepath, smiles_col, name_col, temp_smi)
        if not smi_path:
            return
    else:
        # .smi 文件
        compounds = parse_smi_file(filepath)
        smi_path = str(path)

    # 预览
    print_info(f"检测到 {c(str(len(compounds)), 'g')} 个化合物")
    for i, (sm, nm) in enumerate(compounds[:5]):
        label = nm if nm else sm[:40]
        print(f"     [{i+1}] {label[:55]}")
    if len(compounds) > 5:
        print(f"     ... 共 {len(compounds)} 个")

    # 选择输出目录
    print()
    out_input = input(f"  输出目录 (回车默认 targetnet_results): ").strip().strip('"').strip("'")
    output_dir = get_output_dir(out_input if out_input else None)
    print_ok(f"输出目录: {c(str(output_dir), 'g')}")

    confirm = input(f"\n  直接回车开始预测，输入 'c' 取消: ").strip()
    if confirm.lower() == "c":
        return

    # 提交 (自动分批)
    print_section("提交批量预测")
    records = submit_batch_auto(
        compounds, cfg["fingerprint"], cfg["metric"], cfg["threshold"], output_dir
    )
    if not records:
        print_error("未获取到任何结果")
        return

    # 解析
    print_section("解析结果")
    display_results(records, cfg.get("max_show", 30))

    # 导出到输出目录
    ts = datetime.now().strftime("%Y%m%d")
    ind_dir = output_dir / f"individual_{ts}_results"
    ind_dir.mkdir(parents=True, exist_ok=True)
    base = str(ind_dir / f"targetnet_{path.stem}")

    auto_fmt = cfg.get("auto_export", "csv")
    saved = auto_export(records, base, auto_fmt)
    if saved:
        print_ok("已自动导出:")
        for f in saved:
            print(f"      {c(f, 'g')}")

    # 合并结果表格
    print_section("合并结果")
    merged = merge_results_to_table(records, output_dir)
    if merged:
        print_ok("合并表格已保存:")
        for f in merged:
            print(f"      {c(f, 'g')}")

    # 单个化合物结果
    print_section("单个化合物结果")
    individual = export_individual_results(records, output_dir)
    if individual:
        print_ok(f"已导出 {len(individual)} 个化合物的单独结果 → individual_{ts}_results/")
        print_info(f"位置: {c(str(output_dir / f'individual_{ts}_results'), 'g')}")


def manual_single():
    """手动模式单分子 — 逐步选择参数（保留旧行为）"""
    print_section("手动模式 - 单分子预测")

    print(c("  输入化合物 SMILES + [名称] (空格分隔):", "y"))
    print(c("  输入 'example' 使用示例，空行重新输入", "c"))
    line = ""
    while not line:
        raw = input("  > ").strip()
        if not raw:
            continue
        if raw.lower() == "example":
            line = "C(C=CC1)=C(C=1C(=O)O)O  Salicylic_acid"
            break
        line = raw

    parts = line.split(None, 1)
    smiles = parts[0]
    name = parts[1] if len(parts) > 1 else ""

    fp = _choose_fp()
    mt = _choose_metric()
    th = _input_threshold()

    html = submit_single(smiles, fp, mt, th)
    if not html:
        return

    print_section("解析结果")
    records = parse_single_html(html, smiles, name)
    display_results(records)

    ts = datetime.now().strftime("%Y%m%d")
    default_out = get_output_dir()
    ind_dir = default_out / f"individual_{ts}_results"
    ind_dir.mkdir(parents=True, exist_ok=True)
    label = re.sub(r"[^a-zA-Z0-9_]", "_", name)[:30] if name else "single"
    interactive_export(records, str(ind_dir / f"{label}_targetnet"))


def manual_batch():
    """手动模式批量 — 支持 Excel / .smi，逐步选择参数"""
    print_section("手动模式 - 批量预测")
    print_info("支持格式: Excel (.xlsx/.xls) 或 .smi 文件")
    filepath = input("\n  请输入文件路径: ").strip().strip('"').strip("'")
    if not filepath or not Path(filepath).exists():
        print_error("文件不存在")
        return

    path = Path(filepath)
    smi_path = None
    compounds = None

    if path.suffix.lower() in (".xlsx", ".xls"):
        smiles_col, name_col = interactive_excel_select(filepath)
        if not smiles_col:
            return
        temp_smi = path.with_suffix(".smi")
        smi_path, compounds = excel_to_smi(filepath, smiles_col, name_col, temp_smi)
        if not smi_path:
            return
    else:
        compounds = parse_smi_file(filepath)
        smi_path = str(path)

    print_info(f"检测到 {len(compounds)} 个化合物")

    # 选择输出目录
    out_input = input(f"  输出目录 (回车默认 targetnet_results): ").strip().strip('"').strip("'")
    output_dir = get_output_dir(out_input if out_input else None)

    fp = _choose_fp()
    mt = _choose_metric()
    th = _input_threshold()

    # 提交 (自动分批)
    print_section("提交批量预测")
    records = submit_batch_auto(compounds, fp, mt, th, output_dir)
    if not records:
        print_error("未获取到任何结果")
        return

    print_section("解析结果")
    display_results(records)

    ts = datetime.now().strftime("%Y%m%d")
    ind_dir = output_dir / f"individual_{ts}_results"
    ind_dir.mkdir(parents=True, exist_ok=True)
    interactive_export(records, str(ind_dir / f"targetnet_{path.stem}"))

    # 合并结果表格
    print_section("合并结果")
    merged = merge_results_to_table(records, output_dir)
    if merged:
        print_ok("合并表格已保存:")
        for f in merged:
            print(f"      {c(f, 'g')}")

    # 单个化合物结果
    print_section("单个化合物结果")
    individual = export_individual_results(records, output_dir)
    if individual:
        print_ok(f"已导出 {len(individual)} 个化合物的单独结果 → individual_{ts}_results/")
        print_info(f"位置: {c(str(output_dir / f'individual_{ts}_results'), 'g')}")


def _choose_fp():
    """手动选择指纹"""
    keys = list(FINGERPRINT_MAP.keys())
    print()
    for i, k in enumerate(keys, 1):
        print(f"  [{c(str(i), 'y')}] {FINGERPRINT_LABELS[k]}")
    sel = input(f"\n  选择指纹 [1]: ").strip() or "1"
    try:
        return keys[int(sel) - 1]
    except (ValueError, IndexError):
        return "ecfp4"


def _choose_metric():
    """手动选择指标"""
    keys = list(METRIC_MAP.keys())
    print()
    for i, k in enumerate(keys, 1):
        print(f"  [{c(str(i), 'y')}] {METRIC_LABELS[k]}")
    sel = input(f"\n  选择指标 [1]: ").strip() or "1"
    try:
        return keys[int(sel) - 1]
    except (ValueError, IndexError):
        return "auc"


def _input_threshold():
    while True:
        v = input(f"\n  阈值 (0.1~1.0) [0.7]: ").strip()
        if not v: return 0.7
        try:
            f = float(v)
            if 0.1 <= f <= 1.0: return f
            print_error("0.1~1.0")
        except ValueError:
            print_error("请输入数字")


# ═══════════════════════════════════════════════════════════
#  主菜单
# ═══════════════════════════════════════════════════════════
def main_menu():
    cfg = load_config()
    fp_label = FINGERPRINT_LABELS.get(cfg["fingerprint"], cfg["fingerprint"]).split()[0]
    mt_label = cfg["metric"].upper()
    th_label = cfg["threshold"]

    while True:
        print()
        print(c("  ┌──────────────────────────────────────────────┐", "c"))
        print(c("  │         TargetNet 交互式靶点预测 v3.0         │", "c", bold=True))
        print(c("  ├──────────────────────────────────────────────┤", "c"))
        print(f"  │  当前预设: {c(fp_label, 'g')} | {c(mt_label, 'g')}>={c(str(th_label), 'g')}                │")
        print(c("  ├──────────────────────────────────────────────┤", "c"))
        print(f"  │  {c('[1]', 'y')} ⚡ 快速单分子预测  (一键，用预设参数)      │")
        print(f"  │  {c('[2]', 'y')} ⚡ 快速批量预测    (一键，用预设参数)      │")
        print(f"  │  {c('[3]', 'y')} 📝 手动单分子预测  (逐步选择参数)        │")
        print(f"  │  {c('[4]', 'y')} 📝 手动批量预测    (逐步选择参数)        │")
        print(f"  │  {c('[5]', 'c')} ⚙  预设参数设置    (更改默认指纹/阈值等)  │")
        print(f"  │  {c('[6]', 'c')} ❓ 帮助信息                            │")
        print(f"  │  {c('[0]', 'r')} 退出                                    │")
        print(c("  └──────────────────────────────────────────────┘", "c"))

        choice = input(f"\n  {c('请选择 [0-6]', 'y')}: ").strip()

        if choice == "1":     quick_single()
        elif choice == "2":   quick_batch()
        elif choice == "3":   manual_single()
        elif choice == "4":   manual_batch()
        elif choice == "5":   interactive_settings()
        elif choice == "6":   show_help()
        elif choice == "0":
            print(f"\n  {c('感谢使用，再见！', 'y')}\n")
            break
        else:
            print_error("无效选项，请输入 0-6")


def show_help():
    print_section("帮助信息")
    help_text = """
    TargetNet 是一款基于 QSAR 模型的在线化合物靶点预测工具。

    ⚡ 快速模式 (推荐)
      菜单 [1] [2] — 使用预设参数，只需输入 SMILES 或文件路径即可开始
      预设通过菜单 [5] 配置，支持指纹类型、筛选指标、阈值、导出格式
      一次设置，之后一键预测

    📝 手动模式
      菜单 [3] [4] — 每次预测时逐步选择参数，适合需要临时调整的场景

    【分子指纹】
      ECFP4 — 默认推荐，基于环的拓展连接指纹，药理学最常用
      不确定时使用默认即可

    【筛选阈值】
      阈值越高，模型越可靠，但结果越少。推荐 0.7

    【结果字段说明】
      - Compound_Name : 化合物名称 (可从 .smi 文件名列读取)
      - SMILES        : 化合物 SMILES 结构式
      - UniProt_ID    : 靶点蛋白 UniProt 编号
      - Protein       : 靶点蛋白名称
      - Probability   : 预测结合概率 (0~1)

    【输入文件格式】
      1. Excel (.xlsx) — 自动识别列，选择 SMILES 列和名称列
      2. .smi 文件    — 每行: SMILES [NAME] (空格分隔)
      示例:
        C(C=CC1)=C(C=1C(=O)O)O  Salicylic_acid
        CC(=O)OC1=CC=CC=C1C(=O)O  Aspirin
      名称部分可选，不提供则自动编号 Compound_1, Compound_2...

    【输出目录】
      批量预测结果默认存入 targetnet_results 文件夹
      也可在预测时指定自定义路径
      每次预测自动生成单独的 CSV/Excel + 合并表格

    【相关网站】
      TargetNet:  http://targetnet.scbdd.com/
      SwissTarget: http://www.swisstargetprediction.ch/
      UniProt:   https://www.uniprot.org/
    """
    print(textwrap.dedent(help_text))


def check_deps():
    missing = []
    if not HAS_BS4:   missing.append(("beautifulsoup4", "精确 HTML 解析"))
    if not HAS_PANDAS: missing.append(("pandas", "Excel 导出"))
    if not HAS_COLOR: missing.append(("colorama", "终端彩色"))
    if missing:
        print_warn("部分可选库未安装:")
        for name, reason in missing:
            print(f"      - {name} ({reason})")
        pkgs = " ".join(n for n, _ in missing)
        print_info(f"安装: pip install {pkgs}")
        print()


# ═══════════════════════════════════════════════════════════
#  入口
# ═══════════════════════════════════════════════════════════
if __name__ == "__main__":
    print_banner()
    check_deps()
    try:
        main_menu()
    except KeyboardInterrupt:
        print(f"\n\n  {c('操作已取消', 'y')}\n")
        sys.exit(0)
    except Exception as e:
        print_error(f"异常: {e}")
        import traceback; traceback.print_exc()
        sys.exit(1)
