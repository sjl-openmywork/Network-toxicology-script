"""
SuperPred 靶点预测 — 交互式自动化脚本（一体版）v2.0
===============================================
网站: https://prediction.charite.de/subpages/target_prediction.php
支持: Playwright / Selenium 双引擎，单化合物 / Excel/CSV批量 / 交互菜单

【v2.0 更新】
  - 支持 Excel (.xlsx) 和 CSV 文件输入
  - GUI 弹窗选择 SMILES 列和名称列 (tkinter)
  - 输出默认存入 superpred_results 文件夹（可自定义）
  - 批量预测后自动生成合并表格 + 单个化合物结果

依赖:
  pip install playwright selenium webdriver-manager pandas openpyxl
  playwright install chromium

运行:
  python SuperPred_靶点预测.py              交互模式
  python SuperPred_靶点预测.py --show       交互 + 显示浏览器窗口
  python SuperPred_靶点预测.py demo         快速演示（内置化合物）
  python SuperPred_靶点预测.py file.xlsx    直接用Excel/CSV批量预测（Playwright）
  python SuperPred_靶点预测.py file.csv -e selenium  用Selenium引擎
"""

import os, sys, csv, re, time, json
from datetime import datetime
from pathlib import Path

# ── tkinter (GUI) ──────────────────────────────────────────
try:
    import tkinter as tk
    from tkinter import ttk, messagebox
    HAS_TK = True
except ImportError:
    HAS_TK = False

# ── pandas (Excel支持) ─────────────────────────────────────
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

# ── 全局配置 ──────────────────────────────────────────────
SUPERPRED_URL  = "https://prediction.charite.de/subpages/target_prediction.php"
SCRIPT_DIR     = Path(__file__).parent.absolute()

TIMEOUT_NAVIGATE  = 60       # 秒 — 页面导航
TIMEOUT_ELEMENT   = 15       # 秒 — 元素等待
TIMEOUT_PREDICT   = 300      # 秒 — 预测计算
RETRY_MAX         = 2        # 单化合物最大重试
BATCH_INTERVAL    = 10       # 秒 — 批量间隔（避免限流）

# ── 依赖探测 ──────────────────────────────────────────────
HAS_PLAYWRIGHT = False
HAS_SELENIUM   = False

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError
    HAS_PLAYWRIGHT = True
except ImportError:
    pass

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options as SeleniumOptions
    from selenium.webdriver.chrome.service import Service as SeleniumService
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException as SeleniumTimeout
    from selenium.common.exceptions import (WebDriverException, SessionNotCreatedException)
    from webdriver_manager.chrome import ChromeDriverManager
    HAS_SELENIUM = True
except ImportError:
    pass


# ═══════════════════════════════════════════════════════════
#  引擎抽象层 —— 统一 Playwright / Selenium 接口
# ═══════════════════════════════════════════════════════════

class PlaywrightEngine:
    """Playwright 引擎封装"""
    def __init__(self, headless=True):
        self.headless = headless
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None

    def start(self):
        launch_args = [
            "--no-sandbox", "--disable-setuid-sandbox",
            "--disable-dev-shm-usage", "--ignore-certificate-errors",
            "--disable-blink-features=AutomationControlled",
        ]
        self.playwright = sync_playwright().start()
        
        # 尝试启动浏览器：优先系统 Chrome，其次内置 Chromium
        browser_started = False
        
        # 尝试 1: 系统 Chrome
        try:
            self.browser = self.playwright.chromium.launch(
                channel="chrome", headless=self.headless, args=launch_args)
            print("[OK] 引擎: Playwright + 系统 Chrome")
            browser_started = True
        except Exception as e1:
            # 尝试 2: 内置 Chromium
            try:
                self.browser = self.playwright.chromium.launch(
                    headless=self.headless, args=launch_args)
                print("[OK] 引擎: Playwright + 内置 Chromium")
                browser_started = True
            except Exception as e2:
                print(f"[ERROR] Playwright 启动失败:")
                print(f"  系统 Chrome: {e1}")
                print(f"  内置 Chromium: {e2}")
                print(f"  请运行: playwright install chromium")
                raise RuntimeError("无法启动 Playwright 浏览器")
        
        if not browser_started:
            raise RuntimeError("浏览器启动失败")

        self.context = self.browser.new_context(
            viewport={"width": 1280, "height": 960},
            ignore_https_errors=True,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36",
        )
        self.page = self.context.new_page()
        self.page.set_default_timeout(TIMEOUT_ELEMENT * 1000)

    def stop(self):
        if self.context: self.context.close()
        if self.browser: self.browser.close()
        if self.playwright: self.playwright.stop()

    def goto(self, url):
        self.page.goto(url, wait_until="domcontentloaded", timeout=TIMEOUT_NAVIGATE * 1000)
        try: self.page.wait_for_load_state("networkidle", timeout=30000)
        except PWTimeoutError: pass

    def current_url(self): return self.page.url
    def title(self): return self.page.title()
    def body_text(self): return self.page.locator("body").inner_text()

    def fill(self, selector, text):
        el = self.page.locator(selector)
        if el.count() == 0: return False
        el.wait_for(state="visible", timeout=TIMEOUT_ELEMENT * 1000)
        el.click(); el.fill(""); el.type(text, delay=50)
        return True

    def click(self, selector, wait_for_nav=False):
        el = self.page.locator(selector)
        if el.count() == 0: return False
        el.wait_for(state="visible", timeout=TIMEOUT_ELEMENT * 1000)
        if wait_for_nav:
            el.click()
        else:
            # 使用 JS 点击，不等待导航（避免导航超时）
            el.click(force=True, timeout=5000, no_wait_after=True)
        return True

    def click_nth(self, selector, n, wait_for_nav=False):
        el = self.page.locator(selector)
        if el.count() <= n: return False
        if wait_for_nav:
            el.nth(n).click()
        else:
            el.nth(n).click(force=True, timeout=5000, no_wait_after=True)
        return True

    def wait_url(self, pattern, timeout=None):
        """等待 URL 包含指定模式（轮询方式，避免错过已发生的导航）"""
        if timeout is None:
            timeout = TIMEOUT_PREDICT
        t0 = time.time()
        while time.time() - t0 < timeout:
            try:
                current = self.page.url
                if pattern in current:
                    return True
            except Exception:
                pass
            time.sleep(1)
        raise TimeoutError(f"等待 URL 包含 '{pattern}' 超时 ({timeout}秒)")

    def page_html(self):
        return self.page.content()

    def eval(self, js):
        # Playwright 的 page.evaluate() 把字符串当【表达式】求值，
        # 顶层 return 会报 "Illegal return statement"。
        # 包进 IIFE 函数体，使含 return 的脚本正常工作（与 Selenium 语义一致）。
        wrapped = "(function(){ %s })()" % js
        return self.page.evaluate(wrapped)


class SeleniumEngine:
    """Selenium 引擎封装"""
    def __init__(self, headless=True):
        self.headless = headless
        self.driver = None

    def start(self):
        opts = SeleniumOptions()
        for a in ["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu",
                   "--window-size=1280,960", "--ignore-certificate-errors",
                   "--ignore-ssl-errors", "--disable-blink-features=AutomationControlled"]:
            opts.add_argument(a)
        opts.add_experimental_option("excludeSwitches", ["enable-automation"])
        opts.add_experimental_option("useAutomationExtension", False)
        if self.headless:
            opts.add_argument("--headless=new")

        user_data = str(SCRIPT_DIR / "chrome_user_data")
        os.makedirs(user_data, exist_ok=True)
        opts.add_argument(f"--user-data-dir={user_data}")

        svc = SeleniumService(ChromeDriverManager().install())
        try:
            self.driver = webdriver.Chrome(service=svc, options=opts)
        except SessionNotCreatedException:
            opts.binary_location = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
            self.driver = webdriver.Chrome(service=svc, options=opts)

        self.driver.execute_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        self.driver.set_page_load_timeout(TIMEOUT_PREDICT)
        self.driver.implicitly_wait(5)
        print("✅ 引擎: Selenium + Chrome")

    def stop(self):
        if self.driver: self.driver.quit()

    def goto(self, url):
        self.driver.get(url)
        WebDriverWait(self.driver, TIMEOUT_NAVIGATE).until(
            lambda d: d.execute_script("return document.readyState") == "complete")

    def current_url(self): return self.driver.current_url
    def title(self): return self.driver.title
    def body_text(self): return self.driver.find_element(By.TAG_NAME, "body").text

    def _find(self, css_selector):
        try:
            return WebDriverWait(self.driver, TIMEOUT_ELEMENT).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, css_selector)))
        except SeleniumTimeout:
            return None

    def fill(self, selector, text):
        el = self._find(selector)
        if not el: return False
        el.clear(); el.send_keys(text)
        return True

    def click(self, selector):
        el = self._find(selector)
        if not el: return False
        try: el.click()
        except Exception: self.driver.execute_script("arguments[0].click();", el)
        return True

    def click_nth(self, selector, n):
        els = self.driver.find_elements(By.CSS_SELECTOR, selector)
        if len(els) <= n: return False
        try: els[n].click()
        except Exception: self.driver.execute_script("arguments[0].click();", els[n])
        return True

    def wait_url(self, pattern):
        t0 = time.time()
        while time.time() - t0 < TIMEOUT_PREDICT:
            if pattern in self.driver.current_url: return
            time.sleep(2)

    def page_html(self):
        return self.driver.page_source

    def eval(self, js):
        return self.driver.execute_script(js)


# ═══════════════════════════════════════════════════════════
#  核心业务逻辑（引擎无关）
# ═══════════════════════════════════════════════════════════

def load_csv(csv_path):
    """从CSV加载化合物列表 → [(name, smiles), ...]"""
    compounds = []
    with open(csv_path, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        fieldnames = [fn.lower().strip() for fn in (reader.fieldnames or [])]
        if "smiles" not in fieldnames:
            raise ValueError(f"CSV 缺少 'smiles' 列！当前列: {reader.fieldnames}")
        has_name = "name" in fieldnames
        for i, row in enumerate(reader):
            smiles = row.get("smiles", "").strip()
            if not smiles: continue
            name = row.get("name", "").strip() if has_name else f"Compound_{i+1:03d}"
            if not name: name = f"Compound_{i+1:03d}"
            compounds.append((name, smiles))
    return compounds


def extract_results(engine):
    """定位目标表 → 展开分页 → JS API 全量 → DOM 正则全量"""
    results = []

    # ── 用单次 JS 完成：定位正确的 DataTable + 展开全部行 + 取全量数据 ──
    #   优先 #targets；否则枚举所有 DataTable，选表头含 UniProt/PDB 的（Predicted targets 表）
    js = r"""
    function pickTable() {
        // 1) 首选 id=targets
        try {
            if ($('#targets').length && $.fn.dataTable.isDataTable('#targets'))
                return $('#targets').DataTable();
        } catch(e) {}
        // 2) 枚举所有已初始化的 DataTable，按表头特征选择
        try {
            var settings = $.fn.dataTable.settings || [];
            var fallback = null;
            for (var i = 0; i < settings.length; i++) {
                var api = new $.fn.dataTable.Api(settings[i]);
                var heads = api.columns().header().toArray()
                    .map(function(h){ return (h.textContent||'').trim(); });
                var joined = heads.join('|');
                if (/UniProt|PDB|ChEMBL/i.test(joined) && !/Indication/i.test(joined))
                    return api;                       // Predicted targets 表
                if (fallback === null) fallback = api; // 先记下第一个当兵底
            }
            if (fallback) return fallback;
        } catch(e) {}
        return null;
    }
    try {
        var t = pickTable();
        if (!t) return { error: 'no DataTable found' };
        try { t.page.len(-1).draw(false); } catch(e) {}   // 展开全部行
        return { data: t.rows().data().toArray(),
                 cols: t.columns().header().toArray()
                        .map(function(h){ return (h.textContent||'').trim(); }) };
    } catch(e) { return { error: e.toString() }; }
    """
    try:
        data = engine.eval(js)
        if data and data.get("error"):
            print(f"   ℹ JS API 异常: {data.get('error')}")
        if data and data.get("data") and len(data["data"]) > 0:
            cols = data["cols"]
            for row in data["data"]:
                record = {}
                for i, c in enumerate(cols):
                    val = re.sub(r"<[^>]+>", "", str(row[i])).strip() if i < len(row) else ""
                    record[c] = re.sub(r"\s+", " ", val).strip()
                results.append(record)
            print(f"   [JS API] {len(results)} 行, 列: {cols}")
            return results
    except Exception as e:
        print(f"   ℹ JS API 失败: {e}")

    # B: 正则 HTML（上面已尝试展开分页，此时 tbody 应包含全部行）
    try:
        html = engine.page_html()
        tm = re.search(r'<table[^>]*id=["\']targets["\'][^>]*>(.*?)</table>', html, re.DOTALL | re.I)
        if tm:
            thm = re.search(r"<thead[^>]*>(.*?)</thead>", tm.group(1), re.DOTALL)
            headers = [re.sub(r"<[^>]+>", "", h).strip()
                       for h in re.findall(r"<th[^>]*>(.*?)</th>", thm.group(1), re.DOTALL)] if thm else []
            tbm = re.search(r"<tbody[^>]*>(.*?)</tbody>", tm.group(1), re.DOTALL)
            if tbm:
                for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", tbm.group(1), re.DOTALL):
                    tds = [re.sub(r"<[^>]+>", "", td).strip()
                           for td in re.findall(r"<td[^>]*>(.*?)</td>", tr, re.DOTALL)]
                    if headers and len(tds) >= len(headers):
                        results.append(dict(zip(headers, tds)))
                if results:
                    print(f"   [Regex] {len(results)} 行")
                    return results
    except Exception as e:
        print(f"   ℹ Regex 失败: {e}")

    return results


def save_csv(results, path):
    """保存 CSV"""
    if not results: return
    keys = list(dict.fromkeys(k for r in results for k in r))
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=keys)
        w.writeheader()
        w.writerows(results)
    return path


def make_safe_name(name):
    return re.sub(r'[\\/*?:"<>|]', "_", str(name))


# ═══════════════════════════════════════════════════════════
#  Excel/CSV 读取 + GUI 选列 + 输出管理
# ═══════════════════════════════════════════════════════════

def read_file_columns(filepath):
    """读取 Excel 或 CSV 文件的列名列表"""
    path = Path(filepath)
    try:
        if path.suffix.lower() in (".xlsx", ".xls"):
            if not HAS_PANDAS:
                print("  ❌ 读取 Excel 需要 pandas + openpyxl")
                return None
            df = pd.read_excel(filepath, nrows=0, engine="openpyxl")
            return list(df.columns)
        else:
            if HAS_PANDAS:
                df = pd.read_csv(filepath, nrows=0)
                return list(df.columns)
            else:
                with open(filepath, "r", encoding="utf-8-sig") as f:
                    reader = csv.reader(f)
                    return next(reader, None)
    except Exception as e:
        print(f"  ❌ 读取失败: {e}")
        return None


def load_compounds_from_file(filepath, smiles_col, name_col=None):
    """从 Excel/CSV 加载化合物列表，返回 [(name, smiles), ...]"""
    path = Path(filepath)
    try:
        if path.suffix.lower() in (".xlsx", ".xls"):
            if not HAS_PANDAS:
                print("  ❌ 读取 Excel 需要 pandas + openpyxl")
                return None
            df = pd.read_excel(filepath, engine="openpyxl")
        else:
            if HAS_PANDAS:
                df = pd.read_csv(filepath)
            else:
                compounds = []
                with open(filepath, "r", encoding="utf-8-sig") as f:
                    reader = csv.DictReader(f)
                    for i, row in enumerate(reader):
                        smi = row.get(smiles_col, "").strip()
                        if not smi: continue
                        name = row.get(name_col, "").strip() if name_col else f"Compound_{i+1:03d}"
                        if not name: name = f"Compound_{i+1:03d}"
                        compounds.append((name, smi))
                return compounds
    except Exception as e:
        print(f"  ❌ 读取失败: {e}")
        return None

    if smiles_col not in df.columns:
        print(f"  ❌ SMILES 列 '{smiles_col}' 不存在")
        return None
    if name_col and name_col not in df.columns:
        name_col = None

    df = df.dropna(subset=[smiles_col])
    compounds = []
    for i, (_, row) in enumerate(df.iterrows()):
        smi = str(row[smiles_col]).strip()
        name = str(row[name_col]).strip() if name_col and pd.notna(row.get(name_col)) else f"Compound_{i+1:03d}"
        if not name: name = f"Compound_{i+1:03d}"
        if smi and smi.lower() != "nan":
            compounds.append((name, smi))
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
    root.title("SuperPred - 列选择")
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
    print(f"\n  {'='*50}")
    print(f"  列选择 (终端模式)")
    print(f"  {'='*50}")
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


def get_output_dir(user_specified=None):
    """获取输出目录，默认 superpred_results"""
    if user_specified:
        out = Path(user_specified)
    else:
        out = SCRIPT_DIR / "superpred_results"
    out.mkdir(parents=True, exist_ok=True)
    return out


def save_merged_results(all_records, output_dir):
    """保存合并结果为 XLSX 格式（第1行 Date，第2行 Database，第3行起为数据）"""
    if not all_records:
        return []
    ts = datetime.now().strftime("%Y%m%d")

    all_rows = []
    for name, smiles, results in all_records:
        for r in results:
            row = {"Compound_Name": name, "SMILES": smiles}
            row.update(r)
            all_rows.append(row)

    if not all_rows:
        return []

    if HAS_PANDAS and HAS_OPENPYXL:
        df = pd.DataFrame(all_rows)
        xlsx_path = output_dir / f"superpred_{ts}_merged_results.xlsx"
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["Date", datetime.now().strftime("%Y-%m-%d")])
            ws.append(["Database", "SuperPred"])
            # 写入列标题
            col_labels = list(df.columns)
            ws.append(col_labels)
            # 写入数据行
            for _, row in df.iterrows():
                ws.append(row.tolist())
            wb.save(xlsx_path)
            print(f"  📊 合并Excel: {xlsx_path}")
            return [str(xlsx_path)]
        except Exception as e:
            print(f"  ❌ Excel 导出失败: {e}")
    else:
        print(f"  ⚠ 缺少 pandas/openpyxl，跳过合并 Excel")

    return []


def save_individual_results(all_records, output_dir):
    """按化合物分组导出单独 CSV 结果（第1行 Compound，第2行 SMILES，第3行起为数据）"""
    if not all_records:
        return []

    ts = datetime.now().strftime("%Y%m%d")
    ind_dir = output_dir / f"individual_{ts}_results"
    ind_dir.mkdir(parents=True, exist_ok=True)

    saved = []
    for name, smiles, results in all_records:
        if not results:
            continue
        safe_name = make_safe_name(name)[:50]
        fp = ind_dir / f"{safe_name}_superpred.csv"
        try:
            with open(fp, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                # 第1行：Compound 元数据
                writer.writerow(["Compound", name])
                # 第2行：SMILES 元数据
                writer.writerow(["SMILES", smiles])
                # 第3行：列标题（从第一条结果提取所有键）
                if results:
                    all_keys = list(dict.fromkeys(k for r in results for k in r))
                    writer.writerow(all_keys)
                    # 第4行起：数据
                    for r in results:
                        writer.writerow([r.get(k, "") for k in all_keys])
            saved.append(str(fp))
        except Exception as e:
            print(f"  ❌ {name} 导出失败: {e}")

    return saved


# ═══════════════════════════════════════════════════════════
#  单化合物预测（引擎无关）
# ═══════════════════════════════════════════════════════════

def predict_single(engine, smiles, name=None):
    """对单个化合物执行完整预测流程"""
    label = name or f"Compound({smiles[:20]})"

    header(f"预测: {label}")
    print(f"   SMILES: {smiles}")

    for attempt in range(1, RETRY_MAX + 1):
        if attempt > 1:
            print(f"   ↻ 第 {attempt} 次重试...")

        try:
            # ── Step 1: 打开页面 ──
            step(1, 5, "打开预测页面")
            engine.goto(SUPERPRED_URL)
            time.sleep(2)
            print(f"   {engine.title()}  |  {engine.current_url()}")

            # ── Step 2: 输入 SMILES ──
            step(2, 5, "输入 SMILES")
            ok = engine.fill("#smiles_string", smiles)
            if not ok: ok = engine.fill('input[name="smiles"]', smiles)
            if not ok: raise RuntimeError("未找到 SMILES 输入框")
            time.sleep(0.5)
            print("   ✅ SMILES 已填入")

            # ── Step 3: Search ──
            step(3, 5, "提交 SMILES，加载分子结构")
            clicked = engine.click('input[name="smiles"] + div button[name="start"]')
            if not clicked: clicked = engine.click_nth('button[name="start"]', 1)
            if not clicked: raise RuntimeError("未找到 Search 按钮")

            try:
                # Playwright 的 wait_for_load_state 由 goto 接管;
                # Selenium 需要额外等待
                if isinstance(engine, SeleniumEngine):
                    WebDriverWait(engine.driver, TIMEOUT_NAVIGATE).until(
                        lambda d: d.execute_script("return document.readyState") == "complete")
            except Exception:
                pass
            time.sleep(4)

            # ── Step 4: Start Calculation ──
            step(4, 5, "开始靶点预测计算")
            clicked = engine.click('button[name="searchtype"]')
            if not clicked: clicked = engine.click(
                'form[name="MolForm"] button[type="submit"]')
            if not clicked: raise RuntimeError("未找到 Start Calculation 按钮")

            print("   ⏳ 等待计算结果（1-5 分钟）...")
            t0 = time.time()
            
            # 等待结果页面（轮询，双重检测：URL 变化 或 页面内容出现结果）
            url_changed = False
            while time.time() - t0 < TIMEOUT_PREDICT:
                try:
                    current_url = engine.current_url()
                    # 检查是否跳转到结果页
                    if "target_result" in current_url or "result" in current_url.lower():
                        url_changed = True
                        break
                    # 额外检测：页面内容是否已出现结果表格（URL 未变的情况）
                    try:
                        # 优先检测结果表格 #targets 是否存在且有数据（与 extract_results 一致）
                        has_table = engine.eval(
                            "try{var t=$('#targets').DataTable();"
                            "return t.data().toArray().length>0;}catch(e){"
                            "return document.querySelectorAll('#targets tbody tr').length>0;}")
                        if has_table:
                            print("   ✅ 检测到结果表格（URL 未变）")
                            url_changed = True
                            break
                        page_text = engine.body_text()
                        if ("Known strong binders" in page_text
                                or "Additionally predicted targets" in page_text
                                or "PREDICTION RESULTS" in page_text):
                            print("   ✅ 检测到结果页面内容（URL 未变）")
                            url_changed = True
                            break
                    except Exception:
                        pass
                    # 每 15 秒提示一次进度
                    elapsed = int(time.time() - t0)
                    if elapsed > 0 and elapsed % 15 == 0:
                        print(f"   [{elapsed}s] 计算中... 当前 URL: {current_url[:70]}")
                except Exception:
                    pass
                time.sleep(2)
            
            if not url_changed:
                raise TimeoutError(f"等待结果页面超时 ({TIMEOUT_PREDICT}秒)")
            
            print(f"   ✅ 耗时 {time.time()-t0:.0f} 秒")
            time.sleep(3)  # 等待页面完全加载

            # ── 检查错误 ──
            body = engine.body_text()
            if "No input structure" in body:
                raise RuntimeError("服务器返回: No input structure")

            # ── Step 5: 提取 ──
            step(5, 5, f"提取预测结果 [{engine.current_url()}]")
            results = extract_results(engine)

            if results:
                print(f"\n   ✅ 获取 {len(results)} 条靶点预测")
                key_fields = ["Target", "target", "Name", "Gene", "UniProt",
                              "Probability", "Score", "Model", "ATC", "Drug"]
                for i, r in enumerate(results[:5]):
                    info = {k: r[k] for k in r if any(
                        k.lower().startswith(f.lower()) for f in key_fields)}
                    if not info: info = dict(list(r.items())[:3])
                    print(f"      [{i+1}] {json.dumps(info, ensure_ascii=False)}")
                if len(results) > 5: print(f"      ... 共 {len(results)} 条")
                return results
            else:
                print("   ⚠ 未提取到数据")

        except Exception as e:
            print(f"   [ERROR] {e}")
            # 如果是连接超时，增加等待时间
            if "CONNECTION_TIMED_OUT" in str(e) or "Timeout" in str(e):
                print(f"   检测到连接问题，等待 10 秒后重试...")
                time.sleep(10)

        if attempt < RETRY_MAX:
            time.sleep(8)  # 增加重试间隔

    print(f"   ❌ 全部 {RETRY_MAX} 次尝试均失败")
    return []


def predict_batch(engine, compounds, output_dir=None):
    """批量预测"""
    out_dir = get_output_dir(output_dir)
    out = str(out_dir)
    os.makedirs(out, exist_ok=True)

    # 创建带日期的个体结果子目录（用于即时保存）
    ts_date = datetime.now().strftime("%Y%m%d")
    ind_dir = out_dir / f"individual_{ts_date}_results"
    ind_dir.mkdir(parents=True, exist_ok=True)

    total = len(compounds)
    ok_list, fail_list = [], []

    for i, compound in enumerate(compounds):
        name, smiles = (compound if isinstance(compound, (tuple, list)) and len(compound) == 2
                        else (f"Compound_{i+1:03d}", compound))
        print(f"\n{'#' * 55}\n# [{i+1}/{total}] {name}\n{'#' * 55}")

        results = predict_single(engine, smiles, name)
        if results:
            # 即时保存到带日期的子目录（CSV 格式，3行头）
            safe_name = make_safe_name(name)[:50]
            path = str(ind_dir / f"{i+1:03d}_{safe_name}_superpred.csv")
            with open(path, "w", newline="", encoding="utf-8-sig") as cf:
                cw = csv.writer(cf)
                cw.writerow(["Compound", name])
                cw.writerow(["SMILES", smiles])
                if results:
                    all_keys = list(dict.fromkeys(k for r in results for k in r))
                    cw.writerow(all_keys)
                    for r in results:
                        cw.writerow([r.get(k, "") for k in all_keys])
            print(f"   💾 {path}")
            ok_list.append((name, smiles, results))
        else:
            fail_list.append((name, smiles))

        if i < total - 1:
            print(f"   ⏸ 休息 {BATCH_INTERVAL} 秒...")
            time.sleep(BATCH_INTERVAL)

    # ── 汇总 ──
    print(f"\n{'═' * 55}")
    print(f"📊 完成: {len(ok_list)}/{total} 成功")
    if fail_list: print(f"❌ 失败: {len(fail_list)}")

    # 汇总文件
    summary = []
    summary.append(f"SuperPred 批量预测汇总")
    summary.append(f"时间: {datetime.now():%Y-%m-%d %H:%M:%S}")
    summary.append(f"{'='*45}")
    summary.append(f"\n✅ 成功 ({len(ok_list)}):")
    for n, s, r in ok_list: summary.append(f"  {n}: {len(r)} 靶点")
    if fail_list:
        summary.append(f"\n❌ 失败 ({len(fail_list)}):")
        for n, s in fail_list: summary.append(f"  {n}\t{s}")
    summary_path = os.path.join(out, "_summary.txt")
    with open(summary_path, "w", encoding="utf-8") as f: f.write("\n".join(summary))
    print(f"📄 {summary_path}")

    # 失败CSV
    if fail_list:
        fail_path = os.path.join(out, "_failed_smiles.csv")
        with open(fail_path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f); w.writerow(["name","smiles"]); w.writerows(fail_list)
        print(f"📄 {fail_path}")

    # ── 合并结果 + 单个化合物结果 ──
    print(f"\n{'─' * 55}")
    print(f"📊 导出合并结果...")
    merged = save_merged_results(ok_list, out_dir)
    print(f"📊 导出单个化合物结果...")
    individual = save_individual_results(ok_list, out_dir)
    if individual:
        print(f"  📁 单个结果目录: {out_dir / f'individual_{ts_date}_results'}")

    return ok_list, fail_list


# ═══════════════════════════════════════════════════════════
#  UI 辅助函数
# ═══════════════════════════════════════════════════════════

def header(text):  print(f"\n{'═' * 60}\n  {text}\n{'═' * 60}")
def step(n, total, desc): print(f"\n[{n}/{total}] {desc}...")
def divider(): print("─" * 55)

def banner():
    print("""
╔══════════════════════════════════════════════════════╗
║   SuperPred 靶点预测 — 自动化交互脚本 v2.0           ║
║   prediction.charite.de                              ║
╚══════════════════════════════════════════════════════╝
""")

def show_menu(engine_label):
    divider()
    print(f"  当前引擎: {engine_label}")
    divider()
    print("""
  [1] 🧪 手动输入 SMILES 预测一个化合物
  [2] 📄 从 Excel/CSV 文件批量预测 (GUI选列)
  [3] 🚀 快速演示（5个内置化合物）
  [4] 🔍 测试 SuperPred 网站连通性
  [5] 🔧 切换浏览器引擎
  [6] 👁 切换无头/显示模式
  [7] ⚙  查看当前设置
  [0] 👋 退出
""")

BUILTIN_COMPOUNDS = [
    ("Aspirin",      "CC(=O)OC1=CC=CC=C1C(=O)O"),
    ("Caffeine",     "CN1C=NC2=C1C(=O)N(C(=O)N2C)C"),
    ("Quercetin",    "C1=CC(=C(C=C1C2=C(C(=O)C3=C(C=C(C=C3O2)O)O)O)O)O"),
    ("Resveratrol",  "C1=CC(=CC=C1/C=C/C2=CC(=CC(=C2)O)O)O"),
    ("Curcumin",     "COC1=C(C=CC(=C1)/C=C/C(=O)CC(=O)/C=C/C2=CC(=C(C=C2)O)OC)O"),
]


def check_connectivity(engine):
    header("连通性测试")
    print(f"  目标: {SUPERPRED_URL}")
    try:
        engine.goto(SUPERPRED_URL)
        print(f"  ✅ 连接成功!")
        print(f"  标题: {engine.title()}")
        print(f"  URL:   {engine.current_url()}")
        # 检测关键元素
        html = engine.page_html()
        checks = [
            ("SMILES input", "smiles_string"),
            ("Search button", 'button[name="start"]'),
            ("MolForm", 'name="MolForm"'),
            ("Start Calculation", 'name="searchtype"'),
        ]
        for label, keyword in checks:
            status = "✅" if keyword in html else "⚠"
            print(f"  {status} {label}")
        return True
    except Exception as e:
        print(f"  ❌ 连接失败: {e}")
        print(f"  请检查: 1) 网络  2) VPN  3) 网站是否在线")
        return False


def pick_engine(engine_name=None, headless=True):
    """根据名称创建引擎，失败则尝试另一个"""
    if engine_name == "playwright" and HAS_PLAYWRIGHT:
        try:
            e = PlaywrightEngine(headless=headless)
            e.start()
            return e, "Playwright"
        except Exception as ex:
            print(f"  ⚠ Playwright 启动失败: {ex}")

    if engine_name == "selenium" and HAS_SELENIUM:
        try:
            e = SeleniumEngine(headless=headless)
            e.start()
            return e, "Selenium"
        except Exception as ex:
            print(f"  ⚠ Selenium 启动失败: {ex}")

    # 自动选择
    if engine_name in (None, "auto"):
        for label, has, cls in [
            ("Playwright", HAS_PLAYWRIGHT, PlaywrightEngine),
            ("Selenium",   HAS_SELENIUM,   SeleniumEngine),
        ]:
            if has:
                try:
                    e = cls(headless=headless)
                    e.start()
                    return e, label
                except Exception:
                    continue
    return None, ""


def ensure_deps():
    """检查依赖并给出安装提示"""
    if not HAS_PLAYWRIGHT and not HAS_SELENIUM:
        print("❌ 未检测到 Playwright 或 Selenium！")
        print("   请安装依赖（任选其一）：")
        print("   A) pip install playwright  &&  playwright install chromium")
        print("   B) pip install selenium webdriver-manager   (需要已有Chrome)")
        return False
    return True


# ═══════════════════════════════════════════════════════════
#  交互主循环
# ═══════════════════════════════════════════════════════════

def interactive_mode():
    banner()
    if not ensure_deps(): return

    engine_name = "auto"
    headless = True

    # 首次启动
    engine, label = pick_engine(engine_name, headless)
    if not engine:
        print("❌ 无法启动任何浏览器引擎，请检查依赖安装。")
        return

    try:
        while True:
            show_menu(label)
            choice = input("  请输入选项 [1-7, 0]: ").strip()

            if choice == "0":
                break

            elif choice == "1":
                # ── 手动输入 ──
                header("手动输入化合物")
                smiles = input("  SMILES: ").strip()
                if not smiles:
                    print("  ⚠ SMILES 不能为空")
                    continue
                name = input("  名称（可选）: ").strip() or None
                out_dir = str(SCRIPT_DIR / "superpred_results")
                results = predict_single(engine, smiles, name)
                if results:
                    ts_date = datetime.now().strftime("%Y%m%d")
                    ind_dir = Path(out_dir) / f"individual_{ts_date}_results"
                    ind_dir.mkdir(parents=True, exist_ok=True)
                    fname = make_safe_name(name or f"compound_{smiles[:15]}")
                    path = str(ind_dir / f"{fname}_superpred.csv")
                    with open(path, "w", newline="", encoding="utf-8-sig") as cf:
                        cw = csv.writer(cf)
                        cw.writerow(["Compound", name or ""])
                        cw.writerow(["SMILES", smiles])
                        if results:
                            all_keys = list(dict.fromkeys(k for r in results for k in r))
                            cw.writerow(all_keys)
                            for r in results:
                                cw.writerow([r.get(k, "") for k in all_keys])
                    print(f"   💾 {path}")
                else:
                    print("  ❌ 预测失败")

            elif choice == "2":
                # ── Excel/CSV 批量 ──
                header("Excel/CSV 批量预测")
                file_path = input("  文件路径 (Excel/CSV): ").strip()
                if not os.path.exists(file_path):
                    print(f"  ❌ 文件不存在: {file_path}")
                    continue
                
                ext = Path(file_path).suffix.lower()
                if ext in (".xlsx", ".xls"):
                    # Excel 文件：GUI 选列
                    smiles_col, name_col = gui_excel_select(file_path)
                    if not smiles_col:
                        print("  ❌ 未选择列，取消操作")
                        continue
                    compounds = load_compounds_from_file(file_path, smiles_col, name_col)
                    if not compounds:
                        print("  ❌ 加载化合物失败")
                        continue
                else:
                    # CSV 文件：沿用旧逻辑
                    try:
                        compounds = load_csv(file_path)
                    except Exception as e:
                        print(f"  ❌ CSV 格式错误: {e}")
                        continue
                
                print(f"  ✅ 加载 {len(compounds)} 个化合物")
                if input(f"  确认开始批量预测? [y/N]: ").strip().lower() != "y":
                    continue
                
                ok, fail = predict_batch(
                    engine, compounds,
                    output_dir=None,
                )

            elif choice == "3":
                # ── 快速演示 ──
                header("快速演示")
                print(f"  将预测 {len(BUILTIN_COMPOUNDS)} 个化合物:")
                for i, (n, s) in enumerate(BUILTIN_COMPOUNDS):
                    print(f"    [{i+1}] {n}: {s}")
                if input(f"\n  确认开始? [y/N]: ").strip().lower() != "y":
                    continue
                ok, fail = predict_batch(
                    engine, BUILTIN_COMPOUNDS,
                    output_dir=str(SCRIPT_DIR / "superpred_demo"),
                )

            elif choice == "4":
                # ── 连通性 ──
                check_connectivity(engine)

            elif choice == "5":
                # ── 切换引擎 ──
                opts = []
                if HAS_PLAYWRIGHT: opts.append("playwright")
                if HAS_SELENIUM: opts.append("selenium")
                opts.append("auto")
                print(f"\n  可用引擎: {', '.join(opts)}")
                sel = input(f"  选择 [{opts[0]}]: ").strip().lower()
                if not sel: sel = opts[0]
                if sel in opts:
                    engine.stop()
                    engine, label2 = pick_engine(sel, headless)
                    if engine:
                        label = label2
                    else:
                        print("  ❌ 引擎切换失败，尝试恢复...")
                        engine, label = pick_engine("auto", headless)
                else:
                    print(f"  ❌ 无效选择: {sel}")

            elif choice == "6":
                # ── 切换无头 ──
                headless = not headless
                print(f"\n  当前模式: {'🔒 无头 (后台运行)' if headless else '🖥 显示窗口'}")
                ans = input("  需要重启浏览器生效，立即重启? [y/N]: ").strip().lower()
                if ans == "y":
                    engine.stop()
                    engine, label2 = pick_engine(engine_name, headless)
                    if engine:
                        label = label2
                    else:
                        engine, label = pick_engine("auto", headless)

            elif choice == "7":
                # ── 查看设置 ──
                header("当前设置")
                print(f"  浏览器引擎 : {label}")
                print(f"  显示模式   : {'🖥 显示窗口' if not headless else '🔒 无头'}")
                available = []
                if HAS_PLAYWRIGHT: available.append("Playwright")
                if HAS_SELENIUM: available.append("Selenium")
                print(f"  已安装依赖 : {', '.join(available)}")
                print(f"  最大重试   : {RETRY_MAX} 次/化合物")
                print(f"  批量间隔   : {BATCH_INTERVAL} 秒")
                print(f"  预测超时   : {TIMEOUT_PREDICT} 秒")
                print(f"  默认输出   : {SCRIPT_DIR / 'superpred_results'}")

            else:
                print(f"  ❌ 无效选项: {choice}")

    except KeyboardInterrupt:
        print("\n\n⚠ 用户中断")
    finally:
        engine.stop()
        print("\n👋 已退出")


# ═══════════════════════════════════════════════════════════
#  快速入口
# ═══════════════════════════════════════════════════════════

def cli_entry():
    """命令行直接模式（非交互）"""
    # 解析参数
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    headless = "--show" not in sys.argv
    engine_name = "playwright"  # 默认

    # -e 指定引擎
    for i, a in enumerate(sys.argv[1:]):
        if a == "-e" and i + 2 < len(sys.argv):
            engine_name = sys.argv[i + 2]
            break

    # 查找文件参数
    input_file = None
    for a in args:
        if (a.endswith(".csv") or a.endswith(".xlsx") or a.endswith(".xls")) and os.path.exists(a):
            input_file = a
            break

    if not ensure_deps(): return

    if input_file:
        # 批量模式
        engine, label = pick_engine(engine_name, headless)
        if not engine: return
        try:
            header(f"批量预测: {input_file}  ({label})")
            ext = Path(input_file).suffix.lower()
            if ext in (".xlsx", ".xls"):
                # Excel 文件：自动猜测列（非交互模式下不弹窗）
                columns = read_file_columns(input_file)
                if not columns:
                    print("  ❌ 无法读取文件列名")
                    return
                smi_idx = _guess_smiles_column(columns)
                smiles_col = columns[smi_idx]
                # 尝试猜测名称列
                name_col = None
                name_keywords = ["name", "compound_name", "title", "id", "feature"]
                for col in columns:
                    if col.lower().strip() in name_keywords:
                        name_col = col
                        break
                compounds = load_compounds_from_file(input_file, smiles_col, name_col)
                print(f"  SMILES列: {smiles_col}, 名称列: {name_col or '(无)'}")
            else:
                compounds = load_csv(input_file)
            if not compounds:
                print("  ❌ 加载化合物失败")
                return
            print(f"  共 {len(compounds)} 个化合物")
            predict_batch(
                engine, compounds,
                output_dir=None,
            )
        finally:
            engine.stop()
    else:
        # demo 模式
        banner()
        if not ensure_deps(): return
        engine, label = pick_engine(engine_name, headless)
        if not engine: return
        try:
            header(f"快速演示 ({label})")
            print(f"  共 {len(BUILTIN_COMPOUNDS)} 个化合物")
            predict_batch(
                engine, BUILTIN_COMPOUNDS,
                output_dir=str(SCRIPT_DIR / "superpred_demo"),
            )
        finally:
            engine.stop()


# ═══════════════════════════════════════════════════════════
#  main
# ═══════════════════════════════════════════════════════════

if __name__ == "__main__":
    if len(sys.argv) > 1:
        # 有参数 → 直接模式
        cli_entry()
    else:
        # 无参数 → 交互模式
        interactive_mode()
