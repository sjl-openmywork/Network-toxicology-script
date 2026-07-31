# -*- coding: utf-8 -*-
"""
ChEMBL 化合物靶点检索 — 终端交互版 v2.0
===============================================
原理与网页 search_results 一致：SMILES → 分子 → activity → 靶点汇总。
activity 数据自带 target_pref_name，基因详情可选且并发获取。

【v2.0 更新】
  - 支持 Excel (.xlsx) 和 CSV 文件输入
  - GUI 弹窗选择 SMILES 列和名称列 (tkinter)
  - 输出默认存入 chembl_results 文件夹（可自定义）
  - 批量检索后自动生成合并表格 + 单个化合物结果

直接运行：python ChEMBL_化合物靶点检索.py
"""

import requests
import pandas as pd
import openpyxl
import time
import sys
import os
import re
import csv
from pathlib import Path
from urllib.parse import quote
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

# ─── tkinter (GUI) ──────────────────────────────────────
try:
    import tkinter as tk
    from tkinter import ttk, messagebox
    HAS_TK = True
except ImportError:
    HAS_TK = False

# ═══════════════════ 终端颜色 ═══════════════════
class C:
    RST = "\033[0m";   BOLD = "\033[1m";   DIM = "\033[2m"
    RED = "\033[91m";  GRN = "\033[92m";   YLW = "\033[93m"
    BLU = "\033[94m";  CYN = "\033[96m";   WHT = "\033[97m"
    @staticmethod
    def ok(s):       return f"{C.GRN}{s}{C.RST}"
    @staticmethod
    def warn(s):     return f"{C.YLW}{s}{C.RST}"
    @staticmethod
    def err(s):      return f"{C.RED}{s}{C.RST}"
    @staticmethod
    def title(s):    return f"{C.BOLD}{C.CYN}{s}{C.RST}"
    @staticmethod
    def header(s):   return f"{C.BOLD}{C.BLU}{s}{C.RST}"
    @staticmethod
    def dim(s):      return f"{C.DIM}{s}{C.RST}"
    @staticmethod
    def bold(s):     return f"{C.BOLD}{C.WHT}{s}{C.RST}"


# ═══════════════════ 配置 ═══════════════════
BASE_URL = "https://www.ebi.ac.uk/chembl/api/data"
TIMEOUT, RETRIES, PAGE_SIZE, RATE, WORKERS = 120, 3, 1000, 0.3, 8
SCRIPT_DIR = Path(__file__).resolve().parent
_target_cache = {}
_session = None

def _sess():
    global _session
    if _session is None:
        _session = requests.Session()
        _session.headers = {"User-Agent": "ChEMBL/2.0", "Accept": "application/json"}
    return _session

def api(url, params=None):
    s = _sess()
    for i in range(RETRIES):
        try:
            r = s.get(url, params=params, timeout=TIMEOUT)
            if r.status_code == 200:    return r.json()
            if r.status_code == 429:    time.sleep(15 * (i+1))
            else:                       return None
        except Exception as e:
            time.sleep(3 * (i+1))
    return None


# ═══════════════════ 输入判断 ═══════════════════
def is_smiles(query):
    """判断输入是 SMILES 还是普通名称"""
    smiles_chars = set("=()[]#@%/\\+-.0123456789")
    special = sum(1 for c in query if c in smiles_chars)
    return len(query) > 3 and (special / max(len(query), 1)) > 0.15


# ═══════════════════ 核心 API ═══════════════════
def find_molecule_by_smiles(smiles, search_type="exact"):
    enc = quote(smiles, safe='')
    if search_type == "exact":
        d = api(f"{BASE_URL}/similarity/{enc}/100.json", {"limit": 10})
        if d and d.get("molecules"): return d["molecules"]
        d = api(f"{BASE_URL}/molecule.json",
                {"molecule_structures__canonical_smiles__connectivity": smiles, "limit": 10})
        if d and d.get("molecules"): return d["molecules"]
    elif search_type == "similarity":
        d = api(f"{BASE_URL}/similarity/{enc}/70.json", {"limit": 20})
        if d and d.get("molecules"): return d["molecules"]
    elif search_type == "substructure":
        d = api(f"{BASE_URL}/substructure/{enc}.json", {"limit": 20})
        if d and d.get("molecules"): return d["molecules"]
    return []


def find_molecule_by_name(name, limit=10):
    """通过化合物名称搜索分子（模糊包含）"""
    d = api(f"{BASE_URL}/molecule.json", {"pref_name__icontains": name, "limit": limit})
    if d and d.get("molecules"): return d["molecules"]
    return []


def find_molecule(query, search_type="exact"):
    """自动判断 SMILES / 名称 并搜索"""
    if is_smiles(query):
        print(f"  {C.dim('[识别为 SMILES]')}")
        return find_molecule_by_smiles(query, search_type)
    else:
        print(f"  {C.dim('[识别为化合物名称]')}")
        mols = find_molecule_by_name(query)
        if not mols:
            print(f"  {C.dim('名称无结果, 尝试 SMILES...')}")
            return find_molecule_by_smiles(query, search_type)
        return mols


def get_activities(chembl_id, assay_type="B", pchembl_min=None):
    params = {
        "molecule_chembl_id": chembl_id, "limit": PAGE_SIZE,
        "only": "molecule_chembl_id,target_chembl_id,target_pref_name,"
                "target_organism,target_tax_id,"
                "standard_type,pchembl_value,action_type,assay_type",
    }
    if assay_type:               params["assay_type"] = assay_type
    if pchembl_min is not None:  params["pchembl_value__gte"] = pchembl_min
    all_, off = [], 0
    while True:
        params["offset"] = off
        d = api(f"{BASE_URL}/activity.json", params)
        if not d or "activities" not in d: break
        all_.extend(d["activities"])
        tot = d.get("page_meta", {}).get("total_count", 0)
        off += PAGE_SIZE
        if off >= tot: break
        time.sleep(RATE)
    return all_


def aggregate_targets(acts, name="", query=""):
    tgts = {}
    for a in acts:
        tid = a.get("target_chembl_id", "")
        if not tid: continue
        if tid not in tgts:
            tgts[tid] = {
                "Compound_Name": name, "Input_Query": query,
                "Target_ChEMBL_ID": tid,
                "Target_Name": a.get("target_pref_name", ""),
                "Organism": a.get("target_organism", ""),
                "Tax_ID": a.get("target_tax_id", ""),
                "Target_Type": "",
                "Accessions": "",
                "Gene_Symbol": "", "Uniprot_ID": "",
                "Activity_Count": 0,
            }
        tgts[tid]["Activity_Count"] += 1
    return tgts


def fetch_gene(tid):
    if tid in _target_cache: return _target_cache[tid]
    d = api(f"{BASE_URL}/target/{tid}.json", {
        "only": "target_chembl_id,target_type,target_components"
    })
    gn, acc_list = "", []
    ttype = ""
    if d:
        ttype = d.get("target_type", "")
        for c in d.get("target_components", []):
            acc = c.get("accession", "")
            if acc and acc not in acc_list: acc_list.append(acc)
            for s in c.get("target_component_synonyms", []):
                st = s.get("syn_type", "")
                if st == "GENE_SYMBOL" and not gn: gn = s.get("component_synonym", "")
    info = {
        "gene_name": gn,
        "uniprot_id": acc_list[0] if acc_list else "",
        "target_type": ttype,
        "accessions": "|".join(acc_list),
    }
    _target_cache[tid] = info
    return info


def enrich_gene(targets, show=True):
    if not targets: return
    tids = list(targets.keys())
    done, total = 0, len(tids)
    if show: print(f"    并发获取 {total} 靶点基因详情 (workers={WORKERS})...")
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        fm = {pool.submit(fetch_gene, tid): tid for tid in tids}
        for f in as_completed(fm):
            tid = fm[f]
            try:
                info = f.result()
                targets[tid]["Gene_Symbol"] = info["gene_name"]
                targets[tid]["Uniprot_ID"] = info["uniprot_id"]
                targets[tid]["Target_Type"] = info["target_type"]
                targets[tid]["Accessions"] = info["accessions"]
            except: pass
            done += 1
            if show and done % 10 == 0: print(f"    [{done}/{total}]")
    if show: print(f"    [{total}/{total}] 完成")


# ═══════════════════ 终端 UI ═══════════════════
def cls():          os.system("cls" if os.name=="nt" else "clear")
def enter():        input(f"\n  {C.dim('按 Enter 返回...')}")
def sint(p, lo=None, hi=None):
    while True:
        v = input(p).strip()
        try:
            v = int(v)
            if (lo is None or v>=lo) and (hi is None or v<=hi): return v
        except: pass
        print(C.warn("  ⚠ 无效输入"))

def box(title, w=60):
    print(f"\n  ╔{'═'*(w-4)}╗")
    print(f"  ║ {C.title(title.center(w-6))} ║")
    print(f"  ╚{'═'*(w-4)}╝")

def sep(ch="─", w=66): print(f"  {C.dim(ch*w)}")


def fmt_table(headers, rows):
    if not rows: return "  (无数据)"
    widths = [len(h) for h in headers]
    for r in rows:
        for i, c in enumerate(r): widths[i] = max(widths[i], len(str(c)))
    widths = [min(w, 45) for w in widths]
    s = "─"*(sum(widths)+len(widths)*3+1)
    lines = [f"  ┌{s.replace('─','┬')}┐"]
    hdr = "".join(f"│ {C.bold(h[:w].ljust(w))} " for h, w in zip(headers, widths)) + "│"
    lines.append(f"  {hdr}")
    lines.append(f"  ├{s.replace('─','┼')}┤")
    for r in rows:
        cells = "".join(f"│ {str(c)[:w].ljust(w)} " for c, w in zip(r, widths)) + "│"
        lines.append(f"  {cells}")
    lines.append(f"  └{s.replace('─','┴')}┘")
    return "\n".join(lines)


# ═══════════════════ Excel/CSV 读取与 GUI 选列 ═══════════════════
def read_file_columns(filepath):
    """读取 Excel 或 CSV 文件的列名列表"""
    path = Path(filepath)
    try:
        if path.suffix.lower() in (".xlsx", ".xls"):
            df = pd.read_excel(filepath, nrows=0, engine="openpyxl")
        else:
            df = pd.read_csv(filepath, nrows=0)
        return list(df.columns)
    except Exception as e:
        print(f"\n  {C.err(f'✗ 读取失败: {e}')}")
        return None


def load_compounds_from_file(filepath, smiles_col, name_col=None):
    """从 Excel/CSV 加载化合物列表，返回 [(smiles, name), ...]"""
    path = Path(filepath)
    try:
        if path.suffix.lower() in (".xlsx", ".xls"):
            df = pd.read_excel(filepath, engine="openpyxl")
        else:
            df = pd.read_csv(filepath)
    except Exception as e:
        print(f"\n  {C.err(f'✗ 读取失败: {e}')}")
        return None

    if smiles_col not in df.columns:
        print(f"\n  {C.err(f'✗ SMILES 列 \'{smiles_col}\' 不存在')}")
        return None
    if name_col and name_col not in df.columns:
        name_col = None

    df = df.dropna(subset=[smiles_col])
    compounds = []
    for _, row in df.iterrows():
        smi = str(row[smiles_col]).strip()
        name = str(row[name_col]).strip() if name_col and pd.notna(row.get(name_col)) else ""
        if smi and smi.lower() != "nan":
            compounds.append((smi, name))
    return compounds


def _guess_smiles_column(columns):
    """自动猜测哪一列是 SMILES 列"""
    keywords = ["smiles", "smi", "canonical_smiles", "structure", "mol"]
    for i, col in enumerate(columns):
        if col.lower().strip() in keywords:
            return i
    return 0


def gui_excel_select(filepath):
    """GUI 弹窗选择 Excel/CSV 中的 SMILES 列和名称列"""
    columns = read_file_columns(filepath)
    if columns is None:
        return None, None

    if HAS_TK:
        return _gui_column_picker(filepath, columns)
    else:
        return _terminal_column_picker(filepath, columns)


def _gui_column_picker(filepath, columns):
    """tkinter GUI 弹窗选列"""
    result = {"smiles_col": None, "name_col": None, "confirmed": False}

    root = tk.Tk()
    root.title("ChEMBL - 列选择")
    root.geometry("480x380")
    root.resizable(False, False)
    root.update_idletasks()
    x = (root.winfo_screenwidth() - 480) // 2
    y = (root.winfo_screenheight() - 380) // 2
    root.geometry(f"+{x}+{y}")
    root.attributes("-topmost", True)
    root.focus_force()

    tk.Label(root, text="Excel 列选择",
             font=("Microsoft YaHei UI", 14, "bold")).pack(pady=(18, 4))
    # 作者/站点信息
    info_frame = tk.Frame(root)
    info_frame.pack(pady=(0, 2))
    tk.Label(info_frame, text="author: shenjianlin", font=("Microsoft YaHei UI", 8), fg="#888").pack(side="left", padx=5)
    tk.Label(info_frame, text="site: git@github.com:sjl-openmywork/Network-toxicology-script.git", font=("Microsoft YaHei UI", 8), fg="#888").pack(side="left", padx=5)
    tk.Label(root, text=f"文件: {Path(filepath).name}",
             font=("Microsoft YaHei UI", 9), fg="#555").pack()
    tk.Label(root, text=f"共 {len(columns)} 列",
             font=("Microsoft YaHei UI", 9), fg="#888").pack(pady=(0, 10))

    # SMILES 列
    frame1 = tk.Frame(root)
    frame1.pack(fill="x", padx=30, pady=6)
    tk.Label(frame1, text="SMILES 列:", font=("Microsoft YaHei UI", 10)).pack(side="left")
    smiles_var = tk.StringVar()
    smiles_combo = ttk.Combobox(frame1, textvariable=smiles_var,
                                values=columns, state="readonly", width=28)
    smiles_combo.pack(side="left", padx=(10, 0))
    smiles_combo.current(_guess_smiles_column(columns))

    # 名称列
    frame2 = tk.Frame(root)
    frame2.pack(fill="x", padx=30, pady=6)
    tk.Label(frame2, text="名称列:  ", font=("Microsoft YaHei UI", 10)).pack(side="left")
    name_var = tk.StringVar(value="(不使用)")
    name_combo = ttk.Combobox(frame2, textvariable=name_var,
                              values=["(不使用)"] + list(columns), state="readonly", width=28)
    name_combo.pack(side="left", padx=(10, 0))
    name_combo.current(0)

    # 数据预览
    tk.Label(root, text="数据预览 (前3行):", font=("Microsoft YaHei UI", 9), fg="#666").pack(pady=(12, 2))
    preview_frame = tk.Frame(root)
    preview_frame.pack(fill="both", padx=30, expand=False)
    try:
        ext = Path(filepath).suffix.lower()
        if ext in (".xlsx", ".xls"):
            df_p = pd.read_excel(filepath, nrows=3, engine="openpyxl")
        else:
            df_p = pd.read_csv(filepath, nrows=3)
        tree = ttk.Treeview(preview_frame, columns=list(df_p.columns), show="headings", height=3)
        for cn in df_p.columns:
            tree.heading(cn, text=str(cn))
            tree.column(cn, width=max(80, 400 // len(df_p.columns)))
        for _, row in df_p.iterrows():
            tree.insert("", "end", values=[str(v)[:40] for v in row])
        tree.pack(fill="x")
    except Exception:
        tk.Label(preview_frame, text="(预览失败)", fg="#999").pack()

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
              width=10).pack(side="left", padx=8)
    tk.Button(btn_frame, text="  取消  ", command=on_cancel,
              font=("Microsoft YaHei UI", 10), width=10).pack(side="left", padx=8)

    root.protocol("WM_DELETE_WINDOW", on_cancel)
    root.mainloop()

    if result["confirmed"] and result["smiles_col"]:
        return result["smiles_col"], result["name_col"]
    return None, None


def _terminal_column_picker(filepath, columns):
    """终端降级模式选列"""
    print(f"\n  {C.title('列选择 (终端模式)')}")
    print(f"  文件: {Path(filepath).name} | 共 {len(columns)} 列")
    for i, col in enumerate(columns, 1):
        print(f"    [{i}] {col}")
    sel = input(f"\n  SMILES 列序号 [1]: ").strip() or "1"
    try:
        smi_idx = int(sel) - 1
        if not (0 <= smi_idx < len(columns)): raise ValueError
    except ValueError:
        smi_idx = 0
    smiles_col = columns[smi_idx]

    sel2 = input(f"  名称列序号 (回车跳过): ").strip()
    name_col = None
    if sel2:
        try:
            ni = int(sel2) - 1
            if 0 <= ni < len(columns): name_col = columns[ni]
        except ValueError:
            pass
    return smiles_col, name_col


# ═══════════════════ 输出目录与导出 ═══════════════════
def get_output_dir(user_specified=None):
    """获取输出目录，默认 chembl_results"""
    if user_specified:
        out = Path(user_specified)
    else:
        out = SCRIPT_DIR / "chembl_results"
    out.mkdir(parents=True, exist_ok=True)
    return out


def save_merged_results(all_rows, output_dir):
    """保存合并结果为 XLSX 格式（第1行 Date，第2行 Database，第3行起为数据）"""
    if not all_rows:
        return []
    ts = datetime.now().strftime("%Y%m%d")
    df = pd.DataFrame(all_rows)
    df = df.drop_duplicates(["ChEMBL_ID", "Target_ChEMBL_ID"]).reset_index(drop=True)

    xlsx_path = output_dir / f"chembl_{ts}_merged_results.xlsx"
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Date", datetime.now().strftime("%Y-%m-%d")])
        ws.append(["Database", "ChEMBL"])
        # 写入列标题
        col_labels = list(df.columns)
        ws.append(col_labels)
        # 写入数据行
        for _, row in df.iterrows():
            ws.append(row.tolist())
        wb.save(xlsx_path)
        return [str(xlsx_path)]
    except Exception as e:
        print(f"  {C.err(f'Excel 导出失败: {e}')}")

    return []


def save_individual_results(all_rows, output_dir):
    """按化合物分组导出单独 CSV 结果（第1行 Compound，第2行 SMILES，第3行起为数据）"""
    if not all_rows:
        return []

    ts = datetime.now().strftime("%Y%m%d")
    ind_dir = output_dir / f"individual_{ts}_results"
    ind_dir.mkdir(parents=True, exist_ok=True)

    # 按化合物名分组
    compound_groups = {}
    for r in all_rows:
        name = r.get("Compound_Name", "") or r.get("Input_Query", "") or "Unknown"
        if name not in compound_groups:
            compound_groups[name] = []
        compound_groups[name].append(r)

    saved = []
    for name, rows in compound_groups.items():
        safe_name = re.sub(r'[<>:"/\\|?*]', '_', name)[:50]
        fp = ind_dir / f"{safe_name}_chembl.csv"
        try:
            df = pd.DataFrame(rows)
            df = df.drop_duplicates(["ChEMBL_ID", "Target_ChEMBL_ID"]).reset_index(drop=True)
            smiles = rows[0].get("Input_Query", "") if rows else ""
            with open(fp, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                # 第1行：Compound 元数据
                writer.writerow(["Compound", name])
                # 第2行：SMILES 元数据
                writer.writerow(["SMILES", smiles])
                # 第3行：列标题；第4行起：数据
                writer.writerow(list(df.columns))
                for _, row in df.iterrows():
                    writer.writerow(row.tolist())
            saved.append(str(fp))
        except Exception as e:
            print(f"  {C.err(f'{name} 导出失败: {e}')}")

    return saved


# ═══════════════════ 应用状态 ═══════════════════
class State:
    def __init__(self):
        self.search = "exact"      # exact / similarity / substructure
        self.assay  = "B"          # B / F / None
        self.pch    = None         # pChEMBL 阈值
        self.org    = None         # 物种过滤
        self.nogene = False        # 跳过基因详情
        self.outdir = str(SCRIPT_DIR / "chembl_results")


# ═══════════════════ 检索核心 ═══════════════════
def run_search(state, query, name=""):
    """执行检索并返回结果行列表 (query 可为 SMILES 或名称)"""
    sep(); print(f"  {C.header('化合物')}: {C.bold(name or '(未命名)')}")
    print(f"  {C.header('查询')}: {query[:50]}{'...' if len(query)>50 else ''}")

    # [1/4]
    print(f"\n  {C.CYN}[1/4]{C.RST} 搜索 ChEMBL 分子...")
    mols = find_molecule(query, state.search)
    if not mols: return []
    print(f"  {C.ok(f'✓ 找到 {len(mols)} 个分子')}")

    rows = []
    for mi, mol in enumerate(mols):
        cid = mol.get("molecule_chembl_id", "")
        mname = mol.get("pref_name", "")
        sim = mol.get("similarity", "")
        si = f" {C.dim(f'(相似度: {sim})')}" if sim else ""
        print(f"\n  ── 分子 {mi+1}/{len(mols)}: {C.bold(cid)} ({mname or '未命名'}){si}")

        # [2/4]
        print(f"  {C.CYN}[2/4]{C.RST} 获取活性数据...")
        acts = get_activities(cid, state.assay, state.pch)
        if not acts: print(f"  {C.warn('  ✗ 无活性数据')}"); continue
        print(f"  {C.ok(f'  ✓ {len(acts)} 条记录')}")

        # [3/4]
        print(f"  {C.CYN}[3/4]{C.RST} 汇总靶点...")
        tgts = aggregate_targets(acts, name, query)
        if state.org:
            before = len(tgts)
            tgts = {k:v for k,v in tgts.items() if (v.get("Organism") or "").lower().find(state.org.lower()) >= 0}
            print(f"  {C.ok(f'  ✓ {len(tgts)} 靶点 (人源筛选, 过滤前 {before})')}")
        else:
            print(f"  {C.ok(f'  ✓ {len(tgts)} 个唯一靶点')}")

        # [4/4]
        if not state.nogene and tgts:
            print(f"  {C.CYN}[4/4]{C.RST} 获取基因详情...")
            enrich_gene(tgts)
        else:
            print(f"  {C.CYN}[4/4]{C.RST} 跳过基因详情")

        for tid, ti in tgts.items():
            rows.append({
                "Compound_Name": name, "Input_Query": query,
                "ChEMBL_ID": cid, "ChEMBL_Name": mname,
                "Similarity": sim, **ti,
            })
    return rows


def show_rows(rows, n=50):
    if not rows: return
    r2 = rows[:n]
    h = ["#","ChEMBL_ID","Target_Name","Type","Gene","Uniprot","#Acts"]
    tr = []
    for i, r in enumerate(r2, 1):
        tr.append([str(i), (r.get("ChEMBL_ID") or "")[:12],
                   (r.get("Target_Name") or "")[:24],
                   (r.get("Target_Type") or "")[:10],
                   (r.get("Gene_Symbol") or "")[:8],
                   (r.get("Uniprot_ID") or "")[:8],
                   str(r.get("Activity_Count", "") or "")])
    print(); print(fmt_table(h, tr))
    if len(rows) > n: print(f"\n  {C.dim(f'... 仅显示前 {n} 条, 共 {len(rows)} 条')}")

def save(rows, path, compound_name=""):
    """保存为 CSV 格式（第1行 Compound，第2行 SMILES，第3行起为数据）"""
    if not rows: return
    try:
        df = pd.DataFrame(rows)
        df = df.drop_duplicates(["ChEMBL_ID","Target_ChEMBL_ID"]).reset_index(drop=True)
        os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
        smiles = rows[0].get("Input_Query", "") if rows else ""
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(["Compound", compound_name or rows[0].get("Compound_Name", "")])
            writer.writerow(["SMILES", smiles])
            writer.writerow(list(df.columns))
            for _, row in df.iterrows():
                writer.writerow(row.tolist())
        print(f"\n  {C.ok(f'✓ 已保存: {path}')}  ({len(df)} 条)")
    except Exception as e:
        print(f"\n  {C.err(f'✗ 保存失败: {e}')}")


# ═══════════════════ 菜单 ═══════════════════
def main_menu(state):
    cls(); box("ChEMBL 化合物靶点检索系统 v2.0")
    print(f"""
  {C.bold("当前设置:")}
    {C.dim('├─')} 搜索模式 : {C.GRN}{state.search}{C.RST}{C.dim('  (exact/similarity/substructure)')}
    {C.dim('├─')} 测定类型 : {C.GRN}{state.assay or '不限'}{C.RST}{C.dim('  (B=结合/F=功能/不限)')}
    {C.dim('├─')} 活性阈值 : {C.GRN}{f'pChEMBL >= {state.pch}' if state.pch else '不限'}{C.RST}
    {C.dim('├─')} 物种过滤 : {C.GRN}{state.org or '不限'}{C.RST}
    {C.dim('├─')} 基因详情 : {C.GRN}{'跳过' if state.nogene else '获取'}{C.RST}
    {C.dim('├─')} 靶点缓存 : {C.GRN}{len(_target_cache)}{C.RST} 条
    {C.dim('└─')} 输出目录 : {C.GRN}{state.outdir}{C.RST}
""")
    sep()
    print(f"""
  {C.bold('[1]')} 单化合物检索
  {C.bold('[2]')} 批量检索 (Excel/CSV + GUI选列)
  {C.bold('[3]')} 检索设置
  {C.bold('[4]')} 帮助
  {C.bold('[0]')} 退出
""")
    sep("═")
    return sint(f"  {C.bold('选择 [0-4]')}: ", 0, 4)


def menu_single(state):
    cls(); box("单化合物检索")
    print(f"\n  {C.dim('输入 SMILES 或化合物名称，脚本自动识别:')}")
    print(f"  {C.dim('  SMILES: CC(=O)Oc1ccccc1C(=O)O')}")
    print(f"  {C.dim('  名称:  aspirin / ibuprofen / 二甲双胍')}")
    s = input(f"\n  {C.bold('查询')}: ").strip()
    if not s: print(f"\n  {C.warn('⚠ 输入不能为空')}"); enter(); return
    n = input(f"  {C.bold('标签 (可选)')}: ").strip()
    rows = run_search(state, s, n if n else s)
    if rows:
        show_rows(rows)
        c = input(f"\n  {C.bold('保存 CSV? [y/N]')}: ").strip().lower()
        if c == 'y':
            p = input(f"  {C.bold('路径 (回车=默认)')}: ").strip()
            if not p:
                sn = n.replace(" ","_") if n else "compound"
                ts = datetime.now().strftime("%Y%m%d")
                ind_dir = Path(state.outdir) / f"individual_{ts}_results"
                ind_dir.mkdir(parents=True, exist_ok=True)
                p = str(ind_dir / f"{sn}_chembl.csv")
            save(rows, p, compound_name=(n if n else s))
    else: print(f"\n  {C.warn('未找到靶点')}")
    enter()

def menu_batch(state):
    cls(); box("批量检索 (Excel/CSV)")
    print(f"\n  {C.dim('支持格式: Excel (.xlsx) 或 CSV 文件')}")
    print(f"  {C.dim('将弹窗引导选择 SMILES 列和名称列')}")
    p = input(f"\n  {C.bold('文件路径')}: ").strip().strip('"').strip("'")
    if not os.path.isfile(p): print(f"\n  {C.err(f'✗ 文件不存在')}"); enter(); return

    path = Path(p)

    # GUI/终端 选择列
    smiles_col, name_col = gui_excel_select(p)
    if not smiles_col:
        print(f"\n  {C.warn('⚠ 已取消列选择')}"); enter(); return

    # 加载化合物
    compounds = load_compounds_from_file(p, smiles_col, name_col)
    if not compounds:
        print(f"\n  {C.err('✗ 未提取到有效化合物')}"); enter(); return

    print(f"\n  {C.ok(f'✓ 加载: {len(compounds)} 个化合物')}")
    print(f"  {C.dim('列映射: SMILES→' + smiles_col + (', Name→' + name_col if name_col else ''))}")

    # 预览
    for i, (smi, nm) in enumerate(compounds[:5]):
        label = nm if nm else smi[:40]
        print(f"    [{i+1}] {label[:55]}")
    if len(compounds) > 5:
        print(f"    ... 共 {len(compounds)} 个")

    # 选择输出目录
    print()
    out_input = input(f"  {C.bold('输出目录 (回车默认 chembl_results)')}: ").strip().strip('"').strip("'")
    output_dir = get_output_dir(out_input if out_input else None)
    print(f"  {C.ok(f'✓ 输出目录: {output_dir}')}")

    if input(f"\n  {C.bold('开始检索? [Y/n]')}: ").strip().lower() == 'n': return

    # 逐个检索
    all_rows = []
    for i, (smi, nm) in enumerate(compounds):
        if not smi or smi == "nan": continue
        print(f"\n {C.dim('─'*55)}")
        print(f"  {C.bold(f'[{i+1}/{len(compounds)}]')} {nm or smi[:40]}")
        rows = run_search(state, smi, nm)
        all_rows.extend(rows)
        time.sleep(RATE)

    if not all_rows:
        print(f"\n  {C.warn('无结果')}"); enter(); return

    # 汇总
    rd = pd.DataFrame(all_rows)
    print(f"\n  {C.ok('完成!')} 化合物:{rd['ChEMBL_ID'].nunique()} | "
          f"靶点:{rd['Target_ChEMBL_ID'].nunique()} | 记录:{len(rd)}")

    # 保存合并结果
    print(f"\n  {C.title('合并结果')}")
    merged = save_merged_results(all_rows, output_dir)
    for f in merged:
        print(f"    {C.ok(f)}")

    # 保存单个化合物结果
    print(f"\n  {C.title('单个化合物结果')}")
    individual = save_individual_results(all_rows, output_dir)
    if individual:
        ts = datetime.now().strftime("%Y%m%d")
        print(f"    {C.ok(f'已导出 {len(individual)} 个化合物的单独结果 → individual_{ts}_results/')}")

    if input(f"\n  {C.bold('预览? [y/N]')}: ").strip().lower() == 'y':
        show_rows(all_rows, 20)
    enter()

def menu_settings(state):
    while True:
        cls(); box("检索设置")
        sn={"exact":"精确匹配 (100%)","similarity":"相似性 (70%)","substructure":"子结构搜索"}
        an={"B":"B-结合实验","F":"F-功能实验",None:"不限"}
        units_show = {1:"exact", 2:"similarity", 3:"substructure"}
        assay_show = {1:"B", 2:"F", 3:None}
        print(f"""
  {C.bold("设置:")}
    {C.dim('[1]')} 搜索模式 : {C.GRN}{state.search}{C.RST}  —  {sn.get(state.search,'')}
    {C.dim('[2]')} 测定类型 : {C.GRN}{state.assay or '不限'}{C.RST}  —  {an.get(state.assay,'')}
    {C.dim('[3]')} pChEMBL阈值: {C.GRN}{f'>={state.pch}' if state.pch else '不限'}{C.RST}
    {C.dim('[4]')} 物种过滤 : {C.GRN}{state.org or '不限'}{C.RST}
    {C.dim('[5]')} 基因详情 : {C.GRN}{'跳过 (快)' if state.nogene else '获取'}{C.RST}
    {C.dim('[6]')} 输出目录 : {C.GRN}{state.outdir}{C.RST}
    {C.dim('[7]')} 清空缓存 ({len(_target_cache)} 条)
    {C.dim('[0]')} 返回
""")
        sep()
        c = sint(f"  {C.bold('选择 [0-7]')}: ", 0, 7)
        if c==0: break
        elif c==1:
            print(f"\n  {C.dim('1.exact  2.similarity  3.substructure')}")
            v=sint(f"  {C.bold('选择 [1-3]')}: ",1,3)
            state.search = units_show[v]
            print(C.ok("  ✓"))
        elif c==2:
            print(f"\n  {C.dim('1.B-结合  2.F-功能  3.不限')}")
            v=sint(f"  {C.bold('选择 [1-3]')}: ",1,3)
            state.assay = assay_show[v]
            print(C.ok("  ✓"))
        elif c==3:
            print(f"\n  {C.dim('pChEMBL: 5=10μM  6=1μM  7=100nM  回车=不限')}")
            try: v=input(f"  {C.bold('阈值')}: ").strip(); state.pch=float(v) if v else None; print(C.ok("  ✓"))
            except: print(C.warn("  无效"))
        elif c==4:
            v=input(f"\n  {C.bold('物种 (如 Homo sapiens, 回车=不限)')}: ").strip()
            state.org=v if v else None; print(C.ok("  ✓"))
        elif c==5:
            state.nogene=not state.nogene
            label = "跳过" if state.nogene else "获取"
            print(f"\n  {C.ok(f'✓ 基因详情: {label}')}")
        elif c==6:
            v=input(f"\n  {C.bold('输出目录 (回车=默认 chembl_results)')}: ").strip()
            if not v:
                state.outdir = str(SCRIPT_DIR / "chembl_results")
                os.makedirs(state.outdir, exist_ok=True)
                print(C.ok(f"  ✓ {state.outdir}"))
            else:
                os.makedirs(v, exist_ok=True)
                state.outdir = v
                print(C.ok(f"  ✓ {v}"))
        elif c==7:
            n=len(_target_cache); _target_cache.clear()
            print(f"\n  {C.ok(f'✓ 清空 {n} 条')}")
        if c not in (0,): enter()

def menu_help():
    cls(); box("帮助")
    print(f"""
  {C.bold('ChEMBL 化合物靶点检索 — 使用说明')}

  {C.header('原理')}
  SMILES → ChEMBL分子 → activity数据 → 去重汇总靶点。
  与网页 https://www.ebi.ac.uk/chembl/search_results/... 完全一致。

  {C.header('搜索模式')}
    exact        — 精确结构匹配 (100% 相似度)
    similarity   — 相似性搜索 (70% Tanimoto)
    substructure — 子结构搜索 (含指定子结构)

  {C.header('pChEMBL')}
    pChEMBL = -log₁₀(IC₅₀/EC₅₀)，越大活性越强
    5.0=10μM  6.0=1μM  7.0=100nM

  {C.header('性能')}
    核心数据 (靶点名称) 直接来自 activity 响应，无需额外 API 调用。
    基因详情通过 {WORKERS} 线程并发获取，大幅加速。
    使用 --no-gene-detail 可跳过基因查询，速度最快。
""")
    enter()


# ═══════════════════ 入口 ═══════════════════
def main():
    if os.name=="nt":
        try:
            import ctypes
            ctypes.windll.kernel32.SetConsoleMode(ctypes.windll.kernel32.GetStdHandle(-11),7)
        except: pass
    st = State()
    while True:
        try:
            c = main_menu(st)
            if c==0: cls(); print(f"\n  {C.ok('再见！')}\n"); break
            elif c==1: menu_single(st)
            elif c==2: menu_batch(st)
            elif c==3: menu_settings(st)
            elif c==4: menu_help()
        except KeyboardInterrupt: print(f"\n\n  {C.warn('取消')}"); enter()
        except Exception as e: print(f"\n  {C.err(f'✗ {e}')}"); enter()

if __name__=="__main__": main()
