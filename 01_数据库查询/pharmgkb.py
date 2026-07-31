"""
PharmGKB 基因靶点搜索脚本 v2.1 — 终端交互模式
================================================================================
功能：根据疾病/药物关键词，从 PharmGKB 数据库检索相关基因靶点，导出为 Excel 文件
数据源：https://www.pharmgkb.org (REST API)
API: https://api.pharmgkb.org/v1/site/search

=== 交互菜单 ===
  1. 基因靶点搜索 —— 按疾病/药物关键词检索
  0. 退出

=== v2.1 更新 ===
  无需用户预先指定条数：脚本自动读取 PharmGKB 返回的结果总数，
  并自动翻页抓取全部检索结果后输出。
"""

import os
import sys
import time
import logging
from datetime import datetime

import openpyxl
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ============ 日志配置 ============
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

# ============ 常量配置 ============
BASE_URL = "https://api.pharmgkb.org/v1/site/search"
PAGE_SIZE = 20
REQUEST_TIMEOUT = 30
MAX_RETRIES = 3
RETRY_BACKOFF = 2
RATE_LIMIT_DELAY = 0.5


# ===================================================================
#  API 调用
# ===================================================================

def build_session() -> requests.Session:
    session = requests.Session()
    retry_strategy = Retry(
        total=MAX_RETRIES,
        backoff_factor=RETRY_BACKOFF,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["POST"],
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update({
        "accept": "application/json",
        "accept-language": "zh-CN,zh;q=0.9,en;q=0.8",
        "content-type": "application/json",
        "origin": "https://www.pharmgkb.org",
        "referer": "https://www.pharmgkb.org/",
        "user-agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        ),
        "x-pgkb-website": "production",
    })
    return session


def fetch_page(session: requests.Session, keyword: str, offset: int) -> dict | None:
    payload = {"query": keyword, "connections": [], "objCls": ["gene"], "from": offset}
    try:
        resp = session.post(BASE_URL, json=payload, timeout=REQUEST_TIMEOUT)
        resp.raise_for_status()
        return resp.json()
    except requests.exceptions.Timeout:
        logger.error(f"请求超时 (offset={offset})，已重试 {MAX_RETRIES} 次，跳过此页")
        return None
    except requests.exceptions.ConnectionError:
        logger.error(f"连接错误 (offset={offset})，请检查网络")
        return None
    except requests.exceptions.HTTPError as e:
        logger.error(f"HTTP 错误 (offset={offset}): {e}")
        return None
    except Exception as e:
        logger.error(f"未知错误 (offset={offset}): {type(e).__name__}: {e}")
        return None


def extract_gene_names(data: dict) -> list[str]:
    names = []
    for item in data.get("data", {}).get("hits", []):
        name = item.get("name", "")
        if name:
            names.append(name)
    return names


def get_total_count(data: dict) -> int:
    return data.get("data", {}).get("total", 0)


def search_all_genes(keyword: str) -> tuple[list[str], int]:
    """
    检索关键词对应的全部基因靶点（自动获取所有页）

    先请求首页获取结果总数，再自动翻页抓取全部结果，
    无需用户预先知道或指定条数。

    Returns
    -------
    (unique_genes, total)
        unique_genes : 去重后的基因列表
        total        : PharmGKB 返回的结果总数
    """
    session = build_session()
    all_genes: list[str] = []
    offset = 0

    first_page = fetch_page(session, keyword, offset)
    if first_page is None:
        logger.error("首页请求失败，终止搜索")
        return [], 0

    total = get_total_count(first_page)
    genes = extract_gene_names(first_page)
    all_genes.extend(genes)
    offset += PAGE_SIZE

    # 无结果直接返回
    if total <= 0:
        return list(dict.fromkeys(all_genes)), total

    # 根据总数自动计算总页数，抓取全部结果
    total_pages = (total + PAGE_SIZE - 1) // PAGE_SIZE

    logger.info(f"PharmGKB 共检索到 {total} 条结果，将全部抓取，共 {total_pages} 页")
    logger.info(f"第 1/{total_pages} 页完成，已获取 {len(genes)} 条")

    for page_num in range(2, total_pages + 1):
        time.sleep(RATE_LIMIT_DELAY)
        data = fetch_page(session, keyword, offset)
        if data is None:
            logger.warning(f"第 {page_num}/{total_pages} 页失败，跳过")
            offset += PAGE_SIZE
            continue
        genes = extract_gene_names(data)
        all_genes.extend(genes)
        logger.info(f"第 {page_num}/{total_pages} 页完成，本页 {len(genes)} 条，累计 {len(all_genes)} 条")
        offset += PAGE_SIZE

    unique_genes = list(dict.fromkeys(all_genes))
    logger.info(f"搜索完成：总计 {len(all_genes)} 条（去重后 {len(unique_genes)} 条）")
    return unique_genes, total


# ===================================================================
#  保存
# ===================================================================

DB_NAME = "PharmGKB"


def _output_dir() -> str:
    """输出子文件夹（与脚本同名，如 pharmgkb），不存在时自动创建"""
    script_name = os.path.splitext(os.path.basename(__file__))[0].lower()
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), script_name)
    os.makedirs(out, exist_ok=True)
    return out


def _safe_keyword(keyword: str) -> str:
    """文件名安全处理：移除特殊字符，保留字母数字与词间空格"""
    cleaned = "".join(c for c in keyword if c.isalnum() or c in " _-()（）")
    return " ".join(cleaned.split()) or "output"


def _build_filepath(keyword: str) -> str:
    """标准输出路径：{子文件夹}/{DB}_{检索词}_{YYYYMMDD_HHMMSS}.xlsx"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{DB_NAME}_{_safe_keyword(keyword)}_{timestamp}.xlsx"
    return os.path.join(_output_dir(), filename)


def save_to_excel(genes: list[str], keyword: str) -> str:
    filepath = _build_filepath(keyword)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "基因靶点"
    ws.append(["序号", "基因名称（Gene Symbol）"])
    for idx, gene in enumerate(genes, start=1):
        ws.append([idx, gene])
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 25
    ws.freeze_panes = "A2"

    wb.save(filepath)
    logger.info(f"文件已保存: {filepath}")
    return filepath


# ===================================================================
#  交互菜单
# ===================================================================

def clear_screen():
    os.system("cls" if os.name == "nt" else "clear")


def press_enter():
    input("\n按 Enter 返回主菜单...")


def show_banner():
    clear_screen()
    print("╔" + "═" * 56 + "╗")
    print("║" + "   PharmGKB 基因靶点搜索工具 v2.1".center(52) + "║")
    print("║" + "   数据源: https://www.pharmgkb.org (REST API)".center(48) + "║")
    print("║" + "   author: shenjianlin".center(48) + "║")
    print("║" + "   site: git@github.com:sjl-openmywork/Network-toxicology-script.git".center(36) + "║")
    print("╚" + "═" * 56 + "╝")


def menu_search():
    """菜单 [1] 基因靶点搜索"""
    show_banner()
    print("\n  ▸ 基因靶点搜索")
    print("  " + "─" * 54)
    print("  输入疾病或药物关键词，自动检索并抓取 PharmGKB 全部相关基因靶点\n")

    keyword = input("  请输入关键词（疾病/药物名称，如 Diabetes、Aspirin）: ").strip()
    if not keyword:
        print("\n  ⚠ 关键词不能为空")
        press_enter()
        return

    print()
    genes, total = search_all_genes(keyword)

    if not genes:
        print(f"\n  ✗ 未找到「{keyword}」相关基因靶点")
        press_enter()
        return

    filepath = save_to_excel(genes, keyword)
    print(f"\n  ✓ PharmGKB 共检索到 {total} 条结果")
    print(f"  ✓ 已抓取并去重，共获得 {len(genes)} 个基因靶点")
    print(f"  ✓ 文件: {filepath}")
    print()
    print(f"  前 {min(10, len(genes))} 个基因：")
    for gene in genes[:10]:
        print(f"    • {gene}")
    press_enter()


# ===================================================================
#  主入口
# ===================================================================

def main():
    while True:
        show_banner()
        print()
        print("  ┌────────────────────────────────────────────────────┐")
        print("  │  [1] 基因靶点搜索          疾病/药物 → 基因列表  │")
        print("  │  [0] 退出                                      │")
        print("  └────────────────────────────────────────────────────┘")

        choice = input("\n  请选择 [0-1]: ").strip()

        if choice == "1":
            menu_search()
        elif choice == "0":
            show_banner()
            print("\n  感谢使用 PharmGKB 基因靶点搜索工具，再见！\n")
            break
        else:
            print(f"\n  ⚠ 无效选项「{choice}」，请重新选择")
            time.sleep(1)


if __name__ == "__main__":
    main()
