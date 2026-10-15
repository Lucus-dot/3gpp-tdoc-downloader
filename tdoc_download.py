#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""T0 — 3GPP TDoc 下载工具 (v2)

独立包成员之一(与 tdoc_extract.py / tdoc_direct.py / tdoc_gui.pyw 同目录)。
★ 按号直下(只知道 TDoc 号 / 规范号时直接下载)见同目录 tdoc_direct.py, 本脚本不含该能力。

两用工具, 由 GUI(tdoc_gui.pyw) 或命令行调用:
  用法1 基于 TDoc 清单下载 + 解压 + 按 agenda 嵌套分类(一步到位, 默认输出 by_agenda/ 树);
  用法2 仅分类(对已下载解压好的包按 agenda 分类, 不下载)。分类引擎复用 tdoc_extract.classify。

关键约定/行为:
  - 只读源, 只在 --out 目录新建文件, 绝不修改/删除源清单或共享状态。
  - 下载基址二选一: --base 直给完整目录 URL(推荐, 最可靠); 或 --meeting + --tsg 自动拼。
    ★ 3GPP 目录命名不规整, 强烈建议用 --base 直接粘贴文档目录地址:
        RAN1 → https://www.3gpp.org/ftp/tsg_ran/WG1_RL1/TSGR1_<n>[b]/Docs/
        RAN2 → .../tsg_ran/WG2_RL2/    RAN3 → .../WG3_Iu/    RAN4 → .../WG4_Radio/
        全会 → .../tsg_ran/TSG_RAN/TSGR_<n>/Docs/    (bis 会议目录名末尾用 b, 如 TSGR1_126b)
      若用 --meeting, 请配 --tsg 指明组别(见下); 只给 --meeting 会默认拼成全会的 TSG_RAN, 可能错。
  - HTTP 4xx(403 未授权 / 404 不存在) 视为永久失败, 立即返回不重试; 仅 5xx/网络异常才重试退避。
  - 预检(preflight): 抽样首/中/尾, 若全部 4xx 则判为整会议不可达, 秒级停止(--no-preflight 可关)。
  - 默认分类完成后清理原始 zip 与 extracted/, 仅保留 by_agenda/ 一份; --keep-raw 可保留。
  - --agenda PREFIX 层次化过滤(粗可细): 9.1 命中 9.1/9.1.1/9.1.1.1…, 不会误命中 9.10。

用法:
  # 用法1(推荐: 直接给目录 URL)
  python tdoc_download.py --from-tdoclist tdocList.xlsx --base https://www.3gpp.org/ftp/tsg_ran/WG1_RL1/TSGR1_126b/Docs/ --out Docs
  # 用法1(给会议代码 + 组别)
  python tdoc_download.py --from-tdoclist tdocList.xlsx --tsg RAN1 --meeting TSGR1_126b --out Docs
  # 用法2(仅分类, 不下载)
  python tdoc_download.py --classify --package extracted/ --from-tdoclist tdoclist.csv --out Docs
  # 只检查计划
  python tdoc_download.py --from-tdoclist tdocList.xlsx --base ... --dry-run
"""

import argparse
import csv
import io
import json
import os
import shutil
import re
import sys
import time
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed

import urllib.error
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tdoc_extract import classify, agenda_matches, build_field_map

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

DEFAULT_ROOT = "https://www.3gpp.org/ftp/"
# 组别 -> FTP 目录(WG 目录已按 3GPP 实际结构核对)
TSG_PATH = {
    "RAN":  "tsg_ran/TSG_RAN",     # RAN 全会(TSGR_113 等)
    "RAN1": "tsg_ran/WG1_RL1",     # RAN1
    "RAN2": "tsg_ran/WG2_RL2",     # RAN2
    "RAN3": "tsg_ran/WG3_Iu",      # RAN3
    "RAN4": "tsg_ran/WG4_Radio",   # RAN4
    "SA":   "tsg_sa/TSG_SA",
    "CT":   "tsg_ct/TSG_CT",
    "GERAN": "tsg_geran/TSG_GERAN",
}
HDR = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

# 中文输入法下的全角字符: 看起来一样, 程序按码位比却匹配不上。统一归一成半角。
_FULLWIDTH_MAP = {
    0x3000: " ",   # 全角空格
    0xFF0C: ",",   # ，
    0xFF1B: ";",   # ；
    0x3001: ",",   # 、(顿号)
    0xFF0E: ".",   # ．
    0xFF0F: "/",   # ／
    0xFF1A: ":",   # ：
    0xFF08: "(",   # （
    0xFF09: ")",   # ）
    0xFF0D: "-",   # －
    0xFF0B: "+",   # ＋
    0xFF05: "%",   # ％
    0xFF1D: "=",   # ＝
}
# 列表分隔符: 半角/全角逗号、半角/全角分号、顿号、任意空白(含全角空格)。
# 注: 归一化后全角标点已变半角, 这里再兜一层, 保证"多余空格/混用"都能切开。
DELIM_RE = re.compile(r"[,;\s]+")


def norm_user_input(s):
    """归一化中文输入法常见的全角字符 -> 半角(全角数字/字母/标点/全角空格)。"""
    out = []
    for ch in str(s if s is not None else ""):
        o = ord(ch)
        if 0xFF10 <= o <= 0xFF19 or 0xFF21 <= o <= 0xFF3A or 0xFF41 <= o <= 0xFF5A:
            out.append(chr(o - 0xFEE0))          # 全角 0-9 / A-Z / a-z
        else:
            out.append(_FULLWIDTH_MAP.get(o, ch))
    return "".join(out)


def split_list(s):
    """把用户输入的多值串切成列表: 半角/全角逗号、分号、顿号、空格都行, 连续多个也只算一个。"""
    return [x for x in DELIM_RE.split(norm_user_input(s)) if x]


def build_base(meeting, tsg):
    tsg = norm_user_input(tsg or "RAN").strip().upper()
    meeting = norm_user_input(meeting).strip().strip("/")
    if tsg not in TSG_PATH:
        sys.exit(f"[T0] 未知 tsg: {tsg}, 可选 {list(TSG_PATH)}")
    return f"{DEFAULT_ROOT}{TSG_PATH[tsg]}/{meeting}/Docs/"


# ---------------------------------------------------------------------------
# 会议目录名解析(联网核对, 不靠猜)
# ---------------------------------------------------------------------------
# ★ 实测(2026-10 枚举 3GPP 组目录): 各组 bis/后缀命名 **并不统一**, 所以这里不"拼"
#   目录名, 而是拉取该组真实目录清单, 按会议号匹配出候选再让用户确认。
#     RAN1 (tsg_ran/WG1_RL1)  : TSGR1_92b / TSGR1_126b      <- 小写 b
#     RAN2 (tsg_ran/WG2_RL2)  : TSGR2_91bis / TSGR2_99bis   <- 全拼 bis
#     RAN3 (tsg_ran/WG3_Iu)   : TSGR3_91bis / TSGR3_99bis   <- 全拼 bis
#     RAN4 (tsg_ran/WG4_Radio): TSGR4_94Bis / TSGR4_94_eBis <- 大写 Bis(且带 _e)
#     全会 (tsg_ran/TSG_RAN)  : TSGR_113 / TSGR_88e / TSGR_97e
#     SA   (tsg_sa/TSG_SA)    : TSGS_99_Rotterdam_2023_03   <- 常带城市/日期
#     CT   (tsg_ct/TSG_CT)    : TSGC_99_Rotterdam / CT_106_Madrid
# ---------------------------------------------------------------------------
MEETING_NAMING_HINTS = """会议目录命名实例(实测 3GPP 真实目录, 组别决定子目录):
  RAN1 · 126      -> tsg_ran/WG1_RL1/TSGR1_126
  RAN1 · 126-bis  -> tsg_ran/WG1_RL1/TSGR1_126b      ★ RAN1 的 bis 用【小写 b】
  RAN2 · 91-bis   -> tsg_ran/WG2_RL2/TSGR2_91bis     ★ RAN2 写【全拼 bis】
  RAN3 · 91-bis   -> tsg_ran/WG3_Iu/TSGR3_91bis      ★ RAN3 写【全拼 bis】
  RAN4 · 94-bis   -> tsg_ran/WG4_Radio/TSGR4_94Bis   ★ RAN4 写【大写 Bis】
  全会 RAN · 113  -> tsg_ran/TSG_RAN/TSGR_113
  SA · 99         -> tsg_sa/TSG_SA/TSGS_99_Rotterdam_2023_03  (常带城市/日期)
结论: 后缀写法各组不同(b / bis / Bis / _e / 城市名), 所以请用"生成候选 URL"联网核对,
      或直接粘贴浏览器里打开的那个 Docs 目录地址。"""


def _get_text(url, timeout=45):
    with urllib.request.urlopen(urllib.request.Request(url, headers=HDR), timeout=timeout) as r:
        return r.read().decode("utf-8", "replace")


def _docs_alive(url, timeout=30):
    """Docs 目录是否真实存在。3GPP(IIS) 对不存在的路径返回 403, 存在返回 200。"""
    try:
        with urllib.request.urlopen(urllib.request.Request(url, headers=HDR), timeout=timeout) as r:
            return getattr(r, "status", 200) == 200
    except Exception:
        return False


def list_group_dirs(tsg):
    """列出某组根目录下的全部子目录名(原样大小写)。返回 (root_url, [dir...])。"""
    tsg = (tsg or "RAN").upper()
    if tsg not in TSG_PATH:
        sys.exit(f"[T0] 未知 tsg: {tsg}, 可选 {list(TSG_PATH)}")
    root = f"{DEFAULT_ROOT}{TSG_PATH[tsg]}/"
    html = _get_text(root)
    out, seen = [], set()
    for href, text in re.findall(r'<a\s+href="([^"]+)"[^>]*>([^<]*)</a>', html, re.I):
        t = (text or "").strip().rstrip("/")
        if not t or "/" in t or t.startswith(".") or t in seen:
            continue
        if t.lower().startswith(("www.", "ftp", "sort by")):
            continue
        seen.add(t)
        out.append(t)
    return root, out


def resolve_meeting(tsg, number, verify=True):
    """按会议号在组目录中匹配候选会议文件夹。

    返回 (root_url, [ {folder, has_docs, docs_url}, ... ])，按"像会议目录"排序。
    """
    tsg = norm_user_input(tsg or "RAN").strip().upper()
    num = re.sub(r"[^0-9]", "", norm_user_input(number))   # 全角数字 １２６ 也算
    if not num:
        return f"{DEFAULT_ROOT}{TSG_PATH.get(tsg, '')}/", []
    root, dirs = list_group_dirs(tsg)
    pat = re.compile(r"(?<!\d)" + re.escape(num) + r"(?!\d)")
    cands = [d for d in dirs if pat.search(d)]
    cands.sort(key=lambda d: (not d.upper().startswith("TSG"), len(d), d))
    result = []
    for d in cands:
        docs = f"{DEFAULT_ROOT}{TSG_PATH[tsg]}/{d}/Docs/"
        result.append({"folder": d, "has_docs": _docs_alive(docs) if verify else None,
                       "docs_url": docs})
    return root, result


def recommend_from_filename(path):
    """从 TDoc 清单的**文件名**推荐 (tsg, 会议号); 识别不到就返回 (None, None)。

    仅依据文件名, 不做任何猜测性兜底 —— 文件名里没有会议信息就不推荐。

    例:
      TDoc_List_Meeting_RAN1#126-bis.xlsx -> ("RAN1", "126-bis")
      TDoc_List_Meeting_RAN1#126b.xlsx    -> ("RAN1", "126-bis")
      TDoc_List_Meeting_RAN2#91bis.xlsx   -> ("RAN2", "91-bis")
      TDoc_List_Meeting_TSG_RAN#113.xlsx  -> ("RAN",  "113")
      TDoc_List_Meeting_SA#99.xlsx        -> ("SA",   "99")
      random.xlsx                         -> (None, None)
    """
    stem = norm_user_input(os.path.basename(str(path or "")))   # 兼容全角文件名
    stem = re.sub(r"\.(xlsx|xlsm|xls|csv|zip|txt)$", "", stem, flags=re.I)
    toks = [t for t in re.split(r"[^A-Za-z0-9]+", stem) if t]
    if not toks:
        return None, None
    up = [t.upper() for t in toks]

    tsg = None
    for t in up:
        if re.fullmatch(r"RAN[1-6]", t):
            tsg = t
            break
        if t == "GERAN":
            tsg = "GERAN"
            break
    if tsg is None:
        if any(re.fullmatch(r"TSGR", t) for t in up) or "RAN" in up:
            tsg = "RAN"          # 全会: TSG_RAN / TSGR_113
        elif "SA" in up:
            tsg = "SA"
        elif "CT" in up:
            tsg = "CT"

    num, suf = None, ""
    for i, t in enumerate(up):
        m = re.fullmatch(r"(\d{1,3})([A-Z]{0,3})", t)
        if m:
            num = m.group(1)
            suf = (m.group(2) or "").lower()
            if not suf and i + 1 < len(up) and up[i + 1] in ("BIS", "B"):
                suf = "bis"
            break
    if not num:
        return tsg, None
    if suf in ("bis", "b"):
        return tsg, num + "-bis"
    return tsg, num


def norm_tdoc(line):
    """规整一行 TDoc 号为 '<TDOC>.zip' 形式; 非法返回 None。"""
    s = line.strip()
    if not s:
        return None
    if s.lower().endswith(".zip"):
        s = s[:-4]
    s = os.path.basename(s).strip().strip('"').strip("'")
    if not re.match(r"^[A-Za-z0-9\-_]+$", s):
        return None
    return s + ".zip"


def read_manifest(path):
    out = []
    with open(path, encoding="utf-8") as f:
        for ln in f:
            t = norm_tdoc(ln)
            if t:
                out.append(t)
    return out


# ---------------------------------------------------------------------------
# xlsx 解析(纯 stdlib, 无外部依赖)
# ---------------------------------------------------------------------------
def _col_to_idx(ref):
    """'A1'->0, 'C1'->2, 'AA1'->26"""
    m = re.match(r"([A-Z]+)(\d+)", ref)
    if not m:
        return 0
    idx = 0
    for ch in m.group(1):
        idx = idx * 26 + (ord(ch) - 64)
    return idx - 1


def _xlsx_bytes_to_rows(data):
    """从 xlsx 字节解析为 list[list], 含表头行。支持 shared/inline strings / 数值。"""
    z = zipfile.ZipFile(io.BytesIO(data))
    names = z.namelist()

    shared = []
    if "xl/sharedStrings.xml" in names:
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in root.iter(NS + "si"):
            shared.append("".join(t.text or "" for t in si.iter(NS + "t")))

    sheet = None
    if "xl/worksheets/sheet1.xml" in names:
        sheet = "xl/worksheets/sheet1.xml"
    else:
        for n in names:
            if n.startswith("xl/worksheets/sheet") and n.endswith(".xml"):
                sheet = n
                break
    if sheet is None:
        sys.exit("[T0] xlsx 中未找到 worksheet")

    root = ET.fromstring(z.read(sheet))
    rows = []
    for row in root.iter(NS + "row"):
        cells = {}
        max_idx = -1
        for c in row.findall(NS + "c"):
            ref = c.get("r", "")
            ci = _col_to_idx(ref)
            t = c.get("t")
            v = c.find(NS + "v")
            is_el = c.find(NS + "is")
            val = ""
            if t == "s" and v is not None:
                val = shared[int(v.text)]
            elif t == "inlineStr" and is_el is not None:
                val = "".join(tt.text or "" for tt in is_el.iter(NS + "t"))
            elif v is not None:
                val = v.text
            cells[ci] = val
            max_idx = max(max_idx, ci)
        rows.append([cells.get(i, "") for i in range(max_idx + 1)])
    return rows


def _xlsx_from_zip(zip_path):
    """3GU 门户常把 xlsx 打成 zip: 从 zip 内找首个 xlsx 并解析。"""
    with zipfile.ZipFile(zip_path) as z:
        inner = None
        for n in z.namelist():
            if n.lower().endswith((".xlsx", ".xlsm")):
                inner = n
                break
        if inner is None:
            sys.exit(f"[T0] zip 中未找到 xlsx: {zip_path}")
        return _xlsx_bytes_to_rows(z.read(inner))


# ---------------------------------------------------------------------------
# TDoc 清单(xlsx / csv 自动识别) -> (fieldnames, list[dict])
# ---------------------------------------------------------------------------
def read_tdoclist(path):
    low = path.lower()
    if low.endswith((".xlsx", ".xlsm")):
        with open(path, "rb") as f:
            raw = _xlsx_bytes_to_rows(f.read())
    elif low.endswith(".zip"):
        raw = _xlsx_from_zip(path)
    else:
        with open(path, encoding="utf-8-sig", newline="") as f:
            raw = list(csv.reader(f))
    if not raw:
        sys.exit(f"[T0] TDoc 清单为空: {path}")
    fieldnames = [str(c).strip() for c in raw[0]]
    rows = []
    for r in raw[1:]:
        if len(r) < len(fieldnames):
            r = r + [""] * (len(fieldnames) - len(r))
        rows.append(dict(zip(fieldnames, r)))
    return fieldnames, rows


def find_tdoc_col(fieldnames):
    for col in fieldnames:
        low = (col or "").strip().lower()
        if "tdoc" in low and "list" not in low:
            return col
    return None


def emit_csv_for_t1(fieldnames, rows, out_path):
    """写出归一化 CSV, 列名对齐 T1 的 build_field_map 模糊匹配。"""
    want = []
    for cand in ("Tdoc", "Title", "Source", "Type"):
        for col in fieldnames:
            if col == cand or col.lower() == cand.lower():
                if col not in want:
                    want.append(col)
                break
    agenda_col = None
    for col in fieldnames:
        low = (col or "").strip().lower()
        if ("agenda" in low and "list" not in low and "description" not in low
                and "sort order" not in low):
            agenda_col = col
            break
    if agenda_col and agenda_col not in want:
        want.append(agenda_col)

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(want)
        for r in rows:
            w.writerow([(r.get(c) or "").strip() for c in want])
    return out_path, want


def is_valid_zip(p):
    if not os.path.exists(p) or os.path.getsize(p) == 0:
        return False
    try:
        return zipfile.is_zipfile(p)
    except Exception:
        return False


def need_redownload(p):
    return not is_valid_zip(p)


def fetch(name, base, out_dir, timeout):
    """下载单个 TDoc。4xx 视为永久失败, 立即返回不重试; 5xx/网络异常才重试。"""
    target = os.path.join(out_dir, name)
    part = target + ".part"
    last = ""
    for attempt in range(3):
        try:
            data = urllib.request.urlopen(
                urllib.request.Request(base + name, headers=HDR), timeout=timeout
            ).read()
            with open(part, "wb") as fh:
                fh.write(data)
            os.replace(part, target)
            return name, "ok", len(data)
        except urllib.error.HTTPError as e:
            last = f"HTTP {e.code} {e.reason}"
            if 400 <= e.code < 500:
                # 4xx = 永久性错误(403 未授权 / 404 不存在), 重试无意义
                return name, "FAIL", last
            time.sleep(1.5 * (attempt + 1))
        except Exception as e:
            last = str(e)[:120]
            time.sleep(1.5 * (attempt + 1))
    return name, "FAIL", last


def already_extracted(extracted, base_no_zip):
    if not os.path.isdir(extracted):
        return False
    try:
        return any(base_no_zip in fn for fn in os.listdir(extracted))
    except Exception:
        return False


def _write_index(out_dir, index, package, agenda_filter):
    """写出 tdoc_index.json(分类结果的统一索引)。"""
    out_json = os.path.join(out_dir, "tdoc_index.json")
    with open(out_json, "w", encoding="utf-8") as fh:
        json.dump({
            "tool": "T0",
            "agenda_filter": agenda_filter,
            "package": os.path.abspath(package),
            "tdoc_count": len(index),
            "agendas": sorted({i["agenda"] for i in index}),
            "tdocs": index,
        }, fh, ensure_ascii=False, indent=2)
    classified = sum(1 for i in index if i["file"])
    print(f"[T0] 分类完成: {len(index)} 个 TDoc, 其中 {classified} 个已落盘, "
          f"{len(set(i['agenda'] for i in index))} 个 agenda")
    print(f"[T0] 索引: {out_json}")


def main():
    ap = argparse.ArgumentParser(description="3GPP TDoc 下载与分类工具 (T0)")
    ap.add_argument("--manifest", help="下载清单 txt(每行一个 TDoc 号)")
    ap.add_argument("--from-tdoclist", help="从 3GPP TDoc 清单(xlsx/csv/zip)生成清单")
    ap.add_argument("--classify", action="store_true",
                    help="仅分类模式: 对已有 package 按 agenda 分类(不下载)")
    ap.add_argument("--package", help="[仅分类模式] 已下载解压的 TDoc 文件目录")
    ap.add_argument("--base", help="完整文档目录 URL(以 /Docs/ 结尾, 推荐/最可靠)")
    ap.add_argument("--meeting", help="会议目录名, 如 TSGR_113 / TSGR1_126b")
    ap.add_argument("--tsg", default="RAN",
                    help="组别 RAN/RAN1/RAN2/RAN3/RAN4/SA/CT/GERAN, 默认 RAN(全会)")
    ap.add_argument("--out", default="Docs", help="输出目录, 默认 ./Docs")
    ap.add_argument("--unzip", action="store_true", help="模式A默认解压, 此开关仅为兼容保留")
    ap.add_argument("--keep-raw", action="store_true",
                    help="保留原始 zip 与 extracted/(用于续传/重分类); 默认清理只留 by_agenda/ 一份")
    ap.add_argument("--no-emit-csv", action="store_true",
                    help="--from-tdoclist 时不额外写出 tdoclist_for_T1.csv")
    ap.add_argument("--agenda", default=None,
                    help="只处理该 agenda 子树(可粗可细); 支持多个: 半角或全角逗号/分号/顿号/空格都行, "
                         "如 10.7,9.1 或 10.7，9.1")
    ap.add_argument("--workers", type=int, default=5)
    ap.add_argument("--timeout", type=int, default=90)
    ap.add_argument("--dry-run", action="store_true", help="只打印计划, 不下载")
    ap.add_argument("--no-preflight", action="store_true",
                    help="跳过 4xx 预检(默认抽样首/中/尾, 若全部 4xx 则快速停止)")
    ap.add_argument("--resolve", metavar="NUM",
                    help="按会议号联网匹配真实会议目录, 打印候选 Docs URL 后退出(配合 --tsg)")
    ap.add_argument("--hints", action="store_true",
                    help="打印各组会议目录命名实例后退出")
    args = ap.parse_args()

    # --agenda 支持多个: "10.7,9.1" / "10.7，9.1" / "10.7 9.1" / "10.7、9.1" 都能切
    if args.agenda:
        args.agenda = split_list(args.agenda) or None

    # ===== 会议目录名解析(不下载, 供核对 URL) =====
    if args.hints:
        print(MEETING_NAMING_HINTS)
        return
    if args.resolve:
        tsg_u = str(args.tsg).upper()
        root, cands = resolve_meeting(tsg_u, args.resolve)
        print(f"[T0] 组别 {tsg_u} 根目录: {root}")
        if not cands:
            print(f"[T0] 未在该目录中匹配到含 '{args.resolve}' 的会议文件夹。")
            print(MEETING_NAMING_HINTS)
            sys.exit(2)
        print(f"[T0] 候选 {len(cands)} 个(标签 DOCS=该会议 Docs 目录已存在/可访问):")
        for c in cands:
            tag = "DOCS" if c["has_docs"] else " -- "
            print(f"  [{tag}] {c['folder']}")
            print(f"         {c['docs_url']}")
        return

    # ===== 模式 B: 仅分类(不下载) =====
    if args.classify:
        if not args.package or not args.from_tdoclist:
            sys.exit("[T0] 仅分类模式需要 --package 与 --from-tdoclist(CSV/xlsx 清单)")
        if not os.path.isdir(args.package):
            sys.exit(f"[T0] package 目录不存在: {args.package}")
        fieldnames, rows = read_tdoclist(args.from_tdoclist)
        fm = build_field_map(fieldnames)
        if "agenda" not in fm:
            print("[T0] 警告: 清单未识别到 Agenda 列, 全部归入 unknown", file=sys.stderr)
        os.makedirs(args.out, exist_ok=True)
        index = classify(args.package, rows, args.out, field_map=fm, agenda_filter=args.agenda)
        _write_index(args.out, index, args.package, args.agenda)
        return

    # ===== 模式 A: 下载 + 解压 + 分类 =====
    emit_csv_path = None
    rows = None
    fieldnames = None
    if args.manifest:
        if args.agenda:
            print("[T0] 提示: --manifest 无 agenda 信息, --agenda 过滤将被忽略", file=sys.stderr)
        names = read_manifest(args.manifest)
    elif args.from_tdoclist:
        fieldnames, rows = read_tdoclist(args.from_tdoclist)
        tcol = find_tdoc_col(fieldnames)
        if not tcol:
            sys.exit(f"[T0] TDoc 清单未识别到 TDoc 列: {fieldnames}")
        fm = build_field_map(fieldnames)
        if args.agenda:
            if "agenda" not in fm:
                print("[T0] 警告: 清单无 Agenda 列, --agenda 过滤被忽略", file=sys.stderr)
            else:
                rows = [r for r in rows if agenda_matches(r.get(fm["agenda"], ""), args.agenda)]
        names = []
        for r in rows:
            v = (r.get(tcol) or "").strip()
            if v:
                t = norm_tdoc(v)
                if t:
                    names.append(t)
        if not args.no_emit_csv:
            emit_csv_path = os.path.join(args.out, "tdoclist_for_T1.csv")
            emit_csv_for_t1(fieldnames, rows, emit_csv_path)
    else:
        sys.exit("[T0] 必须指定 --manifest / --from-tdoclist, 或 --classify")

    if not names:
        sys.exit("[T0] 清单为空")

    if args.base:
        base = args.base if args.base.endswith("/") else args.base + "/"
    elif args.meeting:
        base = build_base(args.meeting, args.tsg)
    else:
        sys.exit("[T0] 必须指定 --base 或 --meeting")

    out_dir = args.out
    os.makedirs(out_dir, exist_ok=True)
    extracted = os.path.join(out_dir, "extracted")

    print(f"[T0] 清单 {len(names)} 个 TDoc, base={base}", flush=True)
    if emit_csv_path:
        print(f"[T0] 归一化清单 -> {emit_csv_path} (供 T1 --tdoclist)", flush=True)

    if args.dry_run:
        todo = [n for n in names if need_redownload(os.path.join(out_dir, n))]
        print(f"[T0] dry-run: 需下载 {len(todo)} / 已存在有效 {len(names) - len(todo)}")
        for n in todo[:20]:
            print("   ->", base + n)
        if len(todo) > 20:
            print(f"   ... 其余 {len(todo) - 20} 个省略")
        return

    # ===== 预检(preflight): 抽样首/中/尾, 若全部 4xx 则整会议不可达, 快速停止 =====
    if not args.no_preflight:
        pto = min(args.timeout, 20)

        def _probe_status(url):
            try:
                with urllib.request.urlopen(
                        urllib.request.Request(url, headers=HDR), timeout=pto) as r:
                    return getattr(r, "status", 200)
            except urllib.error.HTTPError as e:
                return e.code
            except Exception as e:
                return str(e)[:80]

        idxs = sorted({0, len(names) // 2, len(names) - 1})
        sample = [names[i] for i in idxs]
        codes = [_probe_status(base + n) for n in sample]
        if codes and all(isinstance(c, int) and 400 <= c < 500 for c in codes):
            print(f"[T0] 预检: 抽样 {sample} 全部返回 {codes} —— 文档目录很可能不对"
                  f"(检查 base/--tsg/--meeting, 如 RAN1 应为 tsg_ran/WG1_RL1/TSGR1_126b/Docs/), "
                  f"或该会议文档未公开。已停止, 不做无谓下载。", flush=True)
            print("[T0] 若确有部分文件可下, 加 --no-preflight 强制继续。", flush=True)
            sys.exit(3)

    log = {}
    permanent = set()   # 4xx 永久失败的文件, 不再重试
    todo = [n for n in names if need_redownload(os.path.join(out_dir, n))]
    print(f"[T0] Pass1: {len(todo)} 需下载 (共 {len(names)})", flush=True)
    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        futs = [ex.submit(fetch, n, base, out_dir, args.timeout) for n in todo]
        done = 0
        for fut in as_completed(futs):
            n, st, sz = fut.result()
            log[n] = (st, sz)
            if st == "FAIL" and str(sz).startswith("HTTP 4"):
                permanent.add(n)
            done += 1
            if done % 50 == 0:
                print(f"  progress {done}/{len(todo)}", flush=True)

    bad = [n for n in names if need_redownload(os.path.join(out_dir, n)) and n not in permanent]
    if bad:
        print(f"[T0] Pass2: {len(bad)} 无效/缺失, 修复中", flush=True)
        with ThreadPoolExecutor(max_workers=args.workers) as ex:
            for fut in as_completed(
                [ex.submit(fetch, n, base, out_dir, args.timeout) for n in bad]
            ):
                n, st, sz = fut.result()
                log[n] = (st, sz)
                if st == "FAIL" and str(sz).startswith("HTTP 4"):
                    permanent.add(n)
    elif permanent:
        print(f"[T0] 跳过 Pass2: {len(permanent)} 个为 4xx 永久失败(如 403/404), 不重试", flush=True)

    ok = sum(1 for n in names if is_valid_zip(os.path.join(out_dir, n)))
    fail = [n for n in names if need_redownload(os.path.join(out_dir, n))]

    # ===== 模式A默认解压(供分类) —— 每个 TDoc 解到自己的子目录, 避免同名互相覆盖 =====
    os.makedirs(extracted, exist_ok=True)
    ue = 0
    for n in names:
        zp = os.path.join(out_dir, n)
        if not is_valid_zip(zp):
            continue
        base_no_zip = n[:-4]
        sub = os.path.join(extracted, base_no_zip)
        if os.path.isdir(sub) and os.listdir(sub):
            ue += 1
            continue
        try:
            with zipfile.ZipFile(zp) as z:
                for m in z.namelist():
                    if m.endswith("/"):
                        continue
                    fn = os.path.basename(m)
                    if not fn:
                        continue
                    os.makedirs(sub, exist_ok=True)
                    dest = os.path.join(sub, fn)
                    with z.open(m) as src, open(dest, "wb") as dfh:
                        dfh.write(src.read())
            ue += 1
        except Exception as e:
            print(f"[T0] 解压失败 {n}: {str(e)[:80]}", file=sys.stderr)
    print(f"[T0] 解压 {ue} 个 TDoc 到 {extracted}", flush=True)

    # ===== 分类(默认输出 by_agenda/) =====
    if rows is not None:
        fm = build_field_map(fieldnames)
        index = classify(extracted, rows, out_dir, field_map=fm, agenda_filter=args.agenda)
    else:
        fake_rows = [{"Tdoc": n[:-4]} for n in names]
        index = classify(extracted, fake_rows, out_dir,
                         field_map={"tdoc": "Tdoc"}, agenda_filter=None)
    _write_index(out_dir, index, extracted, args.agenda)

    # ===== 清理: 未勾选 --keep-raw 时, 逐个删掉"已成功分类的"原件, 只留 by_agenda 一份 =====
    # 旧行为是"一票否决": 只要有 1 个 TDoc 没分类成功就一个都不删, 于是 zip + extracted +
    # by_agenda 三份并存, 与界面上"不勾选=只保留一份 by_agenda"的承诺不符。现改为逐个判定。
    kept_raw = []
    if not args.keep_raw:
        classified = {str(i["tdoc"]).lower() for i in index if i.get("file")}
        rm_zip = rm_dir = 0
        for n in names:
            stem = n[:-4]
            zp = os.path.join(out_dir, n)
            if stem.lower() in classified:
                if os.path.isfile(zp):
                    try:
                        os.remove(zp)
                        rm_zip += 1
                    except OSError:
                        kept_raw.append(n)
                sub = os.path.join(extracted, stem)
                if os.path.isdir(sub):
                    try:
                        shutil.rmtree(sub)
                        rm_dir += 1
                    except OSError:
                        pass
            elif os.path.isfile(zp):
                kept_raw.append(n)
        try:
            if os.path.isdir(extracted) and not os.listdir(extracted):
                os.rmdir(extracted)
        except OSError:
            pass
        print(f"[T0] 已清理 {rm_zip} 个原始 zip 与 {rm_dir} 个 extracted/<TDoc>/"
              f"(这些均已成功落入 by_agenda)", flush=True)
        if kept_raw:
            print(f"[T0] 保留 {len(kept_raw)} 个原件未删(未成功分类, 留作排查): "
                  + ", ".join(kept_raw[:10]) + (" ..." if len(kept_raw) > 10 else ""))
            print("KEPT_RAW:", kept_raw[:20], (" ..." if len(kept_raw) > 20 else ""))
        else:
            print("[T0] 输出目录现只保留 by_agenda/ 一份(另含清单/索引/日志)", flush=True)
        # 旧版曾把内容平铺在 extracted/ 根下, 这里只提示, 不擅自删
        try:
            loose = [x for x in os.listdir(extracted) if os.path.isfile(os.path.join(extracted, x))]
        except OSError:
            loose = []
        if loose:
            print(f"[T0] 提示: {extracted} 下还有 {len(loose)} 个旧版遗留的散落文件(新版按 <TDoc>/ 子目录存放), "
                  f"确认无用可自行删除。")

    log_path = os.path.join(out_dir, "_download_log.csv")
    with open(log_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["tdoc", "status", "size_or_err"])
        for n in names:
            st, sz = log.get(n, ("skip", ""))
            w.writerow([n, st, sz])

    # ===== 会议清单索引: TDoc -> agenda(可多个) / title / source, 供失败明细使用 =====
    meta = {}
    if rows:
        fm2 = build_field_map(fieldnames)
        for r in rows:
            tno = (r.get(fm2.get("tdoc", "")) or "").strip()
            if not tno:
                continue
            d = meta.setdefault(tno.lower(), {"agenda": [], "title": "", "source": ""})
            ag = (r.get(fm2.get("agenda", "")) or "").strip()
            if ag and ag not in d["agenda"]:
                d["agenda"].append(ag)
            if not d["title"]:
                d["title"] = (r.get(fm2.get("title", "")) or "").strip()
            if not d["source"]:
                d["source"] = (r.get(fm2.get("source", "")) or "").strip()

    print(f"[T0] DONE 总={len(names)} 有效={ok} 仍失败={len(fail)} 日志={log_path}")
    print(f"[T0] 分类目录: {os.path.join(out_dir, 'by_agenda')}")
    if fail:
        reasons = Counter()
        for n in fail:
            reason_txt = str(log.get(n, ("", ""))[1] or "")
            key = reason_txt.split(":", 1)[0].strip()
            reasons[key or "unknown"] += 1
        print("[T0] 失败原因统计: " + ", ".join(f"{k}×{v}" for k, v in reasons.most_common()))
        head = fail[:10]
        tail = "" if len(fail) <= 10 else f" ... 其余 {len(fail) - 10} 个省略"
        print("STILL_FAIL:", head, tail)

        # ---- 明细: TDoc 号 / 文件名 / 对应 agenda(可能多个) / 标题 ----
        print(f"[T0] 未取得明细(共 {len(fail)} 个):")
        for n in fail:
            d = meta.get(n[:-4].lower(), {})
            ags = ", ".join(d.get("agenda") or []) or "-"
            ttl = d.get("title") or ""
            if len(ttl) > 60:
                ttl = ttl[:57] + "..."
            print(f"   - {n[:-4]}  [文件名 {n}]  [{str(log.get(n, ('', ''))[1] or '')}]  "
                  f"agenda={ags}" + (f"  标题={ttl}" if ttl else ""))

        fail_csv = os.path.join(out_dir, "_failures.csv")
        try:
            with open(fail_csv, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(["tdoc", "zip_name", "agenda(s)", "reason", "title", "source", "url"])
                for n in fail:
                    d = meta.get(n[:-4].lower(), {})
                    w.writerow([n[:-4], n, " | ".join(d.get("agenda") or []),
                                str(log.get(n, ("", ""))[1] or ""), d.get("title", ""),
                                d.get("source", ""), base + n])
            print(f"[T0] 失败明细已写出: {fail_csv}")
        except OSError:
            pass

        if reasons and all(k.startswith("HTTP 4") for k in reasons):
            print("[T0] 说明: 上述失败全部是 4xx(403 未授权 / 404 不存在), 属个别文件在 3GPP 侧"
                  "不公开或不存在, 不影响其余文件; 本次退出码仍为 0, 不算失败退出。")
        if len(fail) == len(names):
            print("[T0] 提示: 全部下载失败。常见原因: ①文档目录 URL 不对"
                  "(RAN1→tsg_ran/WG1_RL1/TSGR1_<n>b/Docs/; 全会→tsg_ran/TSG_RAN/TSGR_<n>/Docs/); "
                  "②个别/整会议文档受限或未公开; ③网络/代理不可达。")


if __name__ == "__main__":
    main()
