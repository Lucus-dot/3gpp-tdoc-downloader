#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按号直下 —— 3GPP TDoc / TS / TR 按号直接下载（独立包成员）

独立包成员之一(与 tdoc_download.py / tdoc_extract.py / tdoc_gui.pyw 同目录)。
不修改 tdoc_download.py / tdoc_extract.py / tdoc_gui.pyw，只 import 复用其函数。

定位原理（会议不是从号里「解」出来的，是去 3GPP 官方会议表「查」出来的）:
  1. 组 code = **TDoc 号前缀**（R1-2607107 -> R1；RP-261575 -> RP），与工作组无关；
     DynaReport 里 RAN/SA/CT/GERAN 是 91KB 空壳，RP/SP/CP 才有数据。
  2. GET DynaReport/Meetings-<code>.htm -> 每行给出【会议目录名 + TDoc 号区间 + 日期】。
  3. 尾段数字落区间 -> 会议目录；目录与组路径**一律从行内 href 提取，绝不拼接**。
  4. HEAD 校验（3GPP 对不存在路径返回 403）-> fetch() 下载。

四级定位（★兜底 = 指定会议窗口 + 遍历会议）:
  L0 指定会议窗口 : --meeting <会议> -> 在会议表里定位该会议，取【前后各 ±radius 个】
                     目录（radius 默认 5）组成窗口，按号码接近度排序后逐目录 HEAD。
                     --ask 时若前面全落空，会交互式提示用户输入会议号再走 L0。
  L1 区间命中      : 区间按距离排序，HEAD 首个 200（现代会议 1 次命中）
  L2 遍历会议目录  : L1 全落空/未命中 -> 把该前缀的会议目录按「号码距离升序」排好，
                     逐目录 HEAD，命中即停（默认只试前 SCAN_LIMIT=10 个）。
  L3 手工兜底      : --base 直给 Docs 目录 URL（前缀无会议表时用）

配置（可选，优先级 CLI > 配置文件 > 默认值）:
  在脚本目录或当前目录放 tdoc_direct.env（或 .env），键值形如:
      SCAN_LIMIT=10        # L2 遍历上限（默认 10）
      NEIGHBOR_RADIUS=5    # L0 指定会议时前后各扫几个目录（默认 5）
      WORKERS=8            # HEAD/下载并发（默认 8）
      TIMEOUT=90           # 下载超时秒（默认 90）
      OUT=Docs             # 输出根目录（默认 Docs）
      SCAN=1               # 是否启用 L2 遍历（1/0，默认 1）
      CACHE=1              # 索引页（会议表/规范列表）落盘缓存（默认 1）
      CACHE_TTL=604800     # 缓存有效期秒（默认 7 天；0 或负数 = 永不过期）
      CACHE_DIR=           # 缓存目录（留空 = <下载根目录>/.tdoc_cache，即 --out 下）
  模板见同目录 tdoc_direct.env.example。无配置文件则全部用默认值。
  用 --show-config 可查看当前生效配置。

★索引缓存（tdoc_cache.py）:
  会议表 / 规范版本列表 / archive 目录列举会**落盘缓存**（默认 7 天），同日重复运行
  不再重抓（单页可达 400KB）。命中时打印「索引缓存命中 …」。
  **默认放进下载目录**：<out>/.tdoc_cache —— 与下载产物同处，归档/清理一并走。
  **只缓存索引页**：真正的 zip 下载与 HEAD 存在性校验一律实时。
  相关开关：--no-cache / --refresh-cache / --cache-ttl N / --cache-dir <目录> / --clear-cache。

用法:
  python tdoc_direct.py --get-tdoc "R1-2607107, RP-261575" --out Docs
  python tdoc_direct.py --get-spec "38.211, 38.821" --out Docs
  python tdoc_direct.py --direct "R1-2607107, TS 38.211" --out Docs    # 自动识别类型
  python tdoc_direct.py --direct "R1-2607107" --dry-run                # 只定位不下载
  python tdoc_direct.py --get-spec 38.211 --spec-ver i40               # 指定版本(或 v18.4.0)
  python tdoc_direct.py --get-tdoc R1-2607107 --meeting R1-126-bis     # L0 指定会议±5
  python tdoc_direct.py --get-tdoc R1-2607107 --ask                    # 未命中时交互输入会议
  python tdoc_direct.py --get-tdoc R1-2607107 --force-scan             # 强制走 L2 兜底(自测)
  python tdoc_direct.py --get-tdoc R1-2607107 --base <Docs URL>        # L3 手工目录
  python tdoc_direct.py --get-tdoc R1-2607107 --refresh-cache          # 忽略索引缓存重抓
  python tdoc_direct.py --clear-cache                                  # 清空索引缓存

输出:
  <out>/by_tdoc/<TDOC>.zip      TDoc
  <out>/specs/<file>.zip        规范
  <out>/direct_index.json       本次结果索引
  <out>/_direct_log.csv         逐件日志(utf-8-sig)
退出码: 0 正常（个别 4xx 不算失败）/ 2 参数错 / 3 全部失败
"""

import argparse
import csv
import json
import os
import re
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed

import urllib.error
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tdoc_download import (   # noqa: E402
    DEFAULT_ROOT, HDR, _get_text, fetch, is_valid_zip, norm_user_input, split_list,
)
import tdoc_cache as _cache   # noqa: E402  —— 索引页（会议表 / 规范列表）本地磁盘缓存

# 只把**索引页**抓取包上缓存（会议表 / 规范版本列表 / archive 目录列举）。
# 真正的 zip 下载（fetch）与 HEAD 存在性校验（_head_ok）不进缓存，必须实时。
_get_text = _cache.wrap(_get_text)

MEETINGS_URL = "https://www.3gpp.org/DynaReport/Meetings-%s.htm"
SPEC_URL = "https://www.3gpp.org/DynaReport/%s.htm"

# --- 会议表解析（实测：行内链接用**反斜杠** \docs\，不能按 "/docs" 过滤，否则整表滤空）---
ROW_DOCS = re.compile(r"docs[\\/]", re.I)
ROW_FTP = re.compile(r"ftp[\\/]([^\"']+?)[\\/]+docs[\\/]", re.I)
ROW_RANGE = re.compile(r"([A-Z]{1,2}\d?)-(\d{5,7})\s*-\s*([A-Z]{1,2}\d?)-(\d{5,7})")
ROW_LABEL = re.compile(r"([A-Z]{1,2}\d?-\d{2,4}(?:-bis)?)")
ROW_DATE = re.compile(r"(\d{4}-\d{2}-\d{2})\s+(\d{4}-\d{2}-\d{2})")

# --- 规范页解析 ---
SPEC_LINK = re.compile(
    r"(?:https?://www\.3gpp\.org)?(/ftp/Specs/archive/\d+_series/([^/\"]+)/([^/\"]+\.zip))", re.I)
A_HREF = re.compile(r'<a\s+href="([^"]+)"[^>]*>([^<]*)</a>', re.I)

# 规范版本字母 → 主版本号（a=Rel-10 … j=Rel-19）
VER_LETTER = {c: 10 + i for i, c in enumerate("abcdefghijkl")}
VER_RE = re.compile(r"-([a-z])([0-9a-f])([0-9a-f])?$", re.I)

TDOC_IN_RE = re.compile(r"^[A-Za-z]{1,2}\d?-\d{4,7}(?:-\d{2,3})?$")
SPEC_IN_RE = re.compile(r"^\d{2}\.\d{3}$")

# 默认值（配置文件缺失时使用）
DEFAULTS = {
    "SCAN_LIMIT": 10,
    "NEIGHBOR_RADIUS": 5,
    "WORKERS": 8,
    "TIMEOUT": 90,
    "OUT": "Docs",
    "SCAN": 1,
    "CACHE": 1,            # 索引页（会议表 / 规范列表）落盘缓存
    "CACHE_TTL": 604800,   # 缓存有效期秒（7 天）；0 或负数 = 永不过期
}

_TABLE_CACHE = {}          # code -> {"ok", "status", "rows", "n_rows"}
_ENV_CACHE = None          # dict 或 None（未加载）


# ===========================================================================
# 配置（.env / tdoc_direct.env）
# ===========================================================================
def _read_env_file(path):
    data = {}
    try:
        with open(path, "r", encoding="utf-8-sig") as fh:
            for line in fh:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                k = k.strip().upper()
                v = v.strip().strip('"').strip("'")
                if k:
                    data[k] = v
    except Exception:
        pass
    return data


def load_env(path=None):
    """加载配置文件。优先 script 目录，其次当前目录；同名优先 tdoc_direct.env > .env。
    返回 {KEY: str}，并在 '__file__' 键里记录来源路径。"""
    global _ENV_CACHE
    if path:
        _ENV_CACHE = _read_env_file(path)
        _ENV_CACHE["__file__"] = path
        return _ENV_CACHE
    if _ENV_CACHE is not None:
        return _ENV_CACHE
    here = os.path.dirname(os.path.abspath(__file__))
    for d in (here, os.getcwd()):
        for fn in ("tdoc_direct.env", ".env"):
            p = os.path.join(d, fn)
            if os.path.isfile(p):
                _ENV_CACHE = _read_env_file(p)
                _ENV_CACHE["__file__"] = p
                return _ENV_CACHE
    _ENV_CACHE = {}
    return _ENV_CACHE


def cfg_raw(key, cli_val=None):
    """优先级：CLI > 配置文件 > DEFAULTS。"""
    if cli_val is not None:
        return cli_val
    env = load_env()
    if key in env:
        return env[key]
    return DEFAULTS.get(key)


def cfg_int(key, cli_val=None, lo=None):
    v = cfg_raw(key, cli_val)
    try:
        iv = int(str(v).strip())
    except (TypeError, ValueError):
        iv = int(DEFAULTS.get(key, 0))
    if lo is not None and iv < lo:
        iv = lo
    return iv


def cfg_str(key, cli_val=None):
    v = cfg_raw(key, cli_val)
    return str(v) if v is not None else ""


# ===========================================================================
# 基础工具
# ===========================================================================
def tdoc_code(s):
    """'R1-2607107' -> 'R1'；'RP-261575' -> 'RP'。"""
    return str(s).split("-", 1)[0].strip().upper()


def tdoc_num(s):
    """取**尾段数字**：'R1-2607107' -> 2607107；'RP-261575' -> 261575。
    勿用 re.sub(r'\\D','',s)——会把 R2 的 2 也带上，与表内数字不一致。"""
    m = re.search(r"(\d{4,7})\s*$", str(s).strip())
    return int(m.group(1)) if m else None


def tdoc_zip(s):
    s = str(s).strip()
    return s if s.lower().endswith(".zip") else s + ".zip"


def norm_tdoc_input(s):
    """归一化 TDoc 输入：全角→半角、去空格、去 .zip、前缀大写。"""
    t = norm_user_input(s).strip().strip('"').strip("'")
    if t.lower().endswith(".zip"):
        t = t[:-4]
    t = re.sub(r"\s+", "", t)
    return t.upper()


def norm_spec_input(s):
    """'TS 38.211' / 'TR 38.821' / '38211' / '38.211' -> '38.211'。"""
    t = norm_user_input(s).strip().upper()
    t = re.sub(r"^(?:TS|TR|SPEC)\s*", "", t)
    t = re.sub(r"\s+", "", t)
    m = re.match(r"^(\d{2})\.?(\d{3})$", t)
    return "%s.%s" % (m.group(1), m.group(2)) if m else None


def docs_url(fpath):
    return "%s%s/Docs/" % (DEFAULT_ROOT, fpath)


def _norm_fpath(p):
    """折叠重复斜杠、统一正斜杠：'tsg_ct//TSG_CT' -> 'tsg_ct/TSG_CT'。"""
    return re.sub(r"/+", "/", p.replace("\\", "/")).strip("/")


def _head_ok(url, timeout):
    """HEAD 探测；3GPP(IIS) 对不存在路径返回 403，存在返回 200。"""
    try:
        req = urllib.request.Request(url, headers=HDR, method="HEAD")
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return getattr(r, "status", 200) == 200
    except urllib.error.HTTPError:
        return False
    except Exception:
        return False


# ===========================================================================
# 会议表
# ===========================================================================
def load_meeting_table(code):
    """抓 + 解析 Meetings-<code>.htm。同次运行按 code 缓存（索引页可达 900KB）。"""
    if code in _TABLE_CACHE:
        return _TABLE_CACHE[code]
    st, body = None, ""
    try:
        body = _get_text(MEETINGS_URL % code)
        st = 200
    except urllib.error.HTTPError as e:
        st = e.code
    except Exception as e:
        st = str(e)[:80]

    rows, seen = [], set()
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", body, re.S | re.I):
        if not ROW_DOCS.search(tr):
            continue
        text = re.sub(r"<[^>]+>", " ", tr).replace("\u2011", "-").replace("&#8209;", "-")
        text = re.sub(r"\s+", " ", text).strip()
        m = ROW_FTP.search(tr)
        if not m:
            continue
        fpath = _norm_fpath(m.group(1))
        if not fpath or fpath in seen:
            continue
        seen.add(fpath)
        rng, lab, dt = ROW_RANGE.search(text), ROW_LABEL.search(text), ROW_DATE.search(text)
        rows.append({
            "fpath": fpath,
            "label": lab.group(1) if lab else "?",
            "date": dt.group(1) if dt else "?",
            "lo": int(rng.group(2)) if rng else None,
            "hi": int(rng.group(4)) if rng else None,
        })
    info = {"ok": st == 200 and len(rows) > 0, "status": st,
            "rows": rows, "n_rows": len(rows),
            "n_ranged": len([r for r in rows if r["lo"] is not None])}
    _TABLE_CACHE[code] = info
    return info


def _estimate_bounds(rows):
    """给「无区间」行估一个号码区间，供排序用（不做盲扫）。

    会议表按时间倒序排列 → 越往下越旧、号码越小。故对无区间行 i：
      上方(index< i, 更新) 最近的、有区间的行的 lo - 1  = 该行号码上界
      下方(index> i, 更旧) 最近的、有区间的行的 hi + 1  = 该行号码下界
    估不出就留空（排序时排到最后）。
    """
    n = len(rows)
    est = [None] * n
    for i, r in enumerate(rows):
        if r["lo"] is not None:
            est[i] = (r["lo"], r["hi"])
            continue
        hi_e, lo_e = None, None
        for j in range(i - 1, -1, -1):
            if rows[j]["lo"] is not None:
                hi_e = rows[j]["lo"] - 1
                break
        for j in range(i + 1, n):
            if rows[j]["lo"] is not None:
                lo_e = rows[j]["hi"] + 1
                break
        if hi_e is None and lo_e is None:
            est[i] = None
        else:
            est[i] = (lo_e if lo_e is not None else 1,
                      hi_e if hi_e is not None else 10 ** 9)
        if est[i] and est[i][0] > est[i][1]:
            est[i] = (min(est[i]), max(est[i]))
    return est


def order_candidates(rows, digits):
    """候选会议目录排序：命中区间(或估算区间)最前，其次按「号码距离」升序，估不出的最后。"""
    est = _estimate_bounds(rows)

    def key(item):
        i, r = item
        b = est[i]
        if not b:
            return (2, 0, i)
        lo, hi = b
        if lo <= digits <= hi:
            return (0, 0, i)
        return (1, min(abs(digits - lo), abs(digits - hi)), i)
    return [r for _, r in sorted(enumerate(rows), key=key)]


def _meet_key(s):
    """会议标识归一化：'R1-126-bis'/'TSGR1_126b'/'R1-126' -> 可比较串（bis 统一成 B）。"""
    t = re.sub(r"[^A-Za-z0-9]", "", str(s)).upper()
    t = re.sub(r"BIS$", "B", t)
    return t


def _tail_num(s):
    """取字符串里**最后一组**数字（≤3 位）：'R1-126-bis'->126；'RAN1#126bis'->126；'126'->126。"""
    nums = re.findall(r"(\d{1,3})", str(s))
    return int(nums[-1]) if nums else None


def match_meeting_rows(rows, query):
    """把用户给的会议标识（目录名 / label / 序号）匹配到会议表的行。返回行列表。

    匹配顺序（先严后松，避免「裸序号 126」误伤 bis/非 bis 两条）:
      ① 归一化相等（label 或目录名尾段）：R1-126-bis / TSGR1_126b / R1-126
      ② 目录路径精确尾匹配；含字母时才做包含匹配
      ③ 序号退化匹配（带 bis 标志区分）
    """
    q = norm_user_input(str(query)).strip()
    if not q:
        return []
    qu = q.upper()
    qk = _meet_key(qu)

    # ① 归一化相等
    hits = [r for r in rows
            if _meet_key(r["label"]) == qk or _meet_key(r["fpath"].split("/")[-1]) == qk]
    if hits:
        return hits

    # ② 目录路径
    hits = [r for r in rows
            if r["fpath"].upper() == qu or r["fpath"].upper().endswith("/" + qu)]
    if hits:
        return hits
    if re.search(r"[A-Za-z]", qu):
        hits = [r for r in rows if qu in r["fpath"].upper()]
        if hits:
            return hits

    # ③ 序号（±bis）
    n = _tail_num(qu)
    if n is not None:
        want_bis = ("BIS" in qu) or bool(re.search(r"\d+B$", qk))
        hits = [r for r in rows
                if _tail_num(r["label"]) == n
                and ("BIS" in r["label"].upper()) == want_bis]
        if hits:
            return hits
    return []


def neighbour_window(rows, idx, radius):
    """以 idx 为中心、前后各 radius 个行组成的窗口（含自身）。"""
    lo = max(0, idx - radius)
    hi = min(len(rows), idx + radius + 1)
    return rows[lo:hi]


def suggest_meetings(rows, digits, n=5):
    """按「号码接近度」给出前 n 个候选会议（供 --meeting 自动填充）。"""
    est = _estimate_bounds(rows)

    def dist(i):
        b = est[i]
        if not b:
            return 10 ** 9
        lo, hi = b
        if lo <= digits <= hi:
            return 0
        return min(abs(digits - lo), abs(digits - hi))
    idx = sorted(range(len(rows)), key=lambda i: (dist(i), i))[:n]
    return [rows[i] for i in idx]


def probe_dirs(dirs, tdoc_zipname, timeout, workers, limit):
    """逐目录 HEAD 探测，命中即停。返回 (命中的目录, 探测次数)。"""
    if limit is not None and limit <= 0:
        return None, 0
    cand = dirs[:limit] if limit is not None else dirs
    tried, i = 0, 0
    w = max(1, workers)
    while i < len(cand):
        chunk = cand[i:i + w]
        i += w
        with ThreadPoolExecutor(max_workers=len(chunk)) as ex:
            futs = {ex.submit(_head_ok, docs_url(d) + tdoc_zipname, timeout): d for d in chunk}
            for fut in as_completed(futs):
                tried += 1
                if fut.result():
                    for f2 in futs:
                        f2.cancel()
                    return futs[fut], tried
    return None, tried


def resolve_tdoc(tdoc, timeout=25, workers=8, scan=True, scan_limit=10,
                 force_scan=False, base=None, meeting=None, radius=5):
    """定位 TDoc。status ∈ {found, not_found, no_table, unsupported}。"""
    code, digits, z = tdoc_code(tdoc), tdoc_num(tdoc), tdoc_zip(tdoc)
    out = {"tdoc": tdoc, "zip": z, "code": code, "digits": digits,
           "status": None, "dir": None, "docs_base": None, "how": None,
           "tried": 0, "scan_total": None}

    if base:                                   # L3 手工目录
        out.update(status="found", dir="(手工 --base)", docs_base=base, how="manual")
        return out
    if digits is None:
        out["status"] = "unsupported"
        return out

    info = load_meeting_table(code)
    if not info["ok"]:
        out["status"] = "no_table"
        out["note"] = ("Meetings-%s.htm 无可解析会议行(status=%s，空壳页)"
                       % (code, info["status"]))
        return out
    rows = info["rows"]
    out["rows_n"] = len(rows)
    out["n_ranged"] = info["n_ranged"]

    # ---- L0 指定会议 ±radius 窗口（用户已给出会议或经 --ask 输入） ----
    if meeting:
        hits = match_meeting_rows(rows, meeting)
        if not hits:
            out["meeting_note"] = "未在 Meetings-%s 里找到会议 %s" % (code, meeting)
        else:
            idx = rows.index(hits[0])
            win = neighbour_window(rows, idx, radius)
            dirs, seen = [], set()
            for r in order_candidates(win, digits):
                if r["fpath"] not in seen:
                    seen.add(r["fpath"])
                    dirs.append(r["fpath"])
            out["scan_total"] = len(dirs)
            out["scanned"] = len(dirs)
            out["meeting_win"] = ("定位到 %s，取前后各 %d 个 → %d 个目录"
                                  % (hits[0]["label"], radius, len(dirs)))
            d, tried = probe_dirs(dirs, z, timeout, workers, len(dirs))
            out["tried"] += tried
            if d:
                out.update(status="found", dir=d, docs_base=docs_url(d), how="meeting")
                return out

    ordered = order_candidates(rows, digits)

    # ---- L1 区间命中（只用「真区间」的行，最可靠） ----
    if not force_scan:
        hits = [r for r in ordered if r["lo"] is not None and r["lo"] <= digits <= r["hi"]]
        if hits:
            d, tried = probe_dirs([r["fpath"] for r in hits], z, timeout, workers, 8)
            out["tried"] += tried
            if d:
                out.update(status="found", dir=d, docs_base=docs_url(d), how="range")
                return out

    # ---- L2 兜底：遍历该前缀的会议目录（受 scan_limit 限制） ----
    if scan or force_scan:
        dirs, seen = [], set()
        for r in ordered:
            if r["fpath"] not in seen:
                seen.add(r["fpath"])
                dirs.append(r["fpath"])
        out["scan_total"] = len(dirs)
        d, tried = probe_dirs(dirs, z, timeout, workers, scan_limit)
        out["tried"] += tried
        out["scanned"] = min(len(dirs), scan_limit)
        if d:
            out.update(status="found", dir=d, docs_base=docs_url(d),
                       how="scan(forced)" if force_scan else "scan")
            return out

    out["status"] = "not_found"
    out["suggest"] = suggest_meetings(rows, digits, 5)
    return out


# ===========================================================================
# 规范（TS / TR）
# ===========================================================================
def parse_ver(fn):
    """'38211-j50.zip' -> (19,5,0)；'38211-ia0.zip' -> (18,10,0)。失败返回 None。"""
    stem = fn[:-4] if fn.lower().endswith(".zip") else fn
    m = VER_RE.search(stem)
    if not m:
        return None
    major = VER_LETTER.get(m.group(1).lower())
    if major is None:
        return None
    return (major, int(m.group(2), 16), int(m.group(3), 16) if m.group(3) else 0)


def ver_str(v):
    return "v%d.%d.%d" % v if v else "?"


def fetch_spec_versions(spec):
    """索引 A：DynaReport/<digits>.htm。返回 (versions, source)，按版本降序。"""
    digits = spec.replace(".", "")
    series = spec.split(".")[0]
    try:
        html = _get_text(SPEC_URL % digits)
    except Exception:
        html = ""

    seen, out = set(), []
    for path, sub, fn in SPEC_LINK.findall(html):
        if sub.lower() != spec.lower() or fn.lower() in seen:
            continue
        seen.add(fn.lower())
        out.append({"file": fn, "url": "https://www.3gpp.org" + path, "ver": parse_ver(fn)})
    if out:
        if all(v["ver"] for v in out):
            out.sort(key=lambda v: v["ver"], reverse=True)   # 只按解析出的版本号排
        return out, "dynareport"

    # 兜底：直接列举 archive 目录（顺序为字典序，这里按版本号重排）
    url = "%sSpecs/archive/%s_series/%s/" % (DEFAULT_ROOT, series, spec)
    try:
        html = _get_text(url)
    except Exception:
        return [], "none"
    out2 = []
    for href, text in A_HREF.findall(html):
        t = (text or "").strip()
        if not t.lower().endswith(".zip") or t.lower() in seen:
            continue
        seen.add(t.lower())
        out2.append({"file": t, "url": url + t, "ver": parse_ver(t)})
    out2.sort(key=lambda v: (v["ver"] is None, [-x for x in v["ver"]] if v["ver"] else []))
    return out2, "ftp_listing"


def pick_version(versions, want):
    """选版本。want: None=最新 / 'i40' / 'j50' / 'v18.4.0' / '18.4.0'。返回 (ver, err)。"""
    if not versions:
        return None, "该规范未解析到任何版本"
    if not want:
        return versions[0], None
    w = norm_user_input(want).strip()
    if re.fullmatch(r"[A-Za-z][0-9A-Fa-f]{1,2}", w):                 # 文件后缀，如 i40
        for v in versions:
            if re.search(r"-%s\.zip$" % re.escape(w), v["file"], re.I):
                return v, None
        return None, "未找到版本后缀 %s" % w
    m = re.match(r"^v?(\d+)\.(\d+)\.(\d+)$", w, re.I)                # 版本号，如 v18.4.0
    if m:
        tgt = (int(m.group(1)), int(m.group(2)), int(m.group(3)))
        for v in versions:
            if v["ver"] == tgt:
                return v, None
        avail = ", ".join(ver_str(v["ver"]) for v in versions[:6])
        return None, "未找到 %s（前几个可用：%s）" % (want, avail)
    return None, "无法识别的版本写法：%s" % want


# ===========================================================================
# 下载与主流程
# ===========================================================================
def download_one(name, base, out_dir, timeout):
    os.makedirs(out_dir, exist_ok=True)
    n, st, sz = fetch(name, base, out_dir, timeout)
    ok = st == "ok" and is_valid_zip(os.path.join(out_dir, n))
    return {"zip": n, "status": "ok" if ok else "FAIL",
            "size_or_err": sz, "file": os.path.join(out_dir, n) if ok else None}


def _print_suggest(sug):
    if not sug:
        return
    print("     可能的会议（按号码接近度，可直接用作 --meeting 的值）：", flush=True)
    for r in sug:
        print("       - %-14s %-10s %s"
              % (r["label"], r["date"], r["fpath"].split("/")[-1]), flush=True)


def _print_not_found(r, args):
    if r.get("meeting_note"):
        why = r["meeting_note"] + "（该模式未走全表遍历）"
    elif r.get("scanned"):
        why = "已遍历 %s / %s 个会议目录" % (r.get("scanned"), r.get("scan_total"))
    elif not args.scan and not args.force_scan:
        why = "已按区间探测 %s 个目录（L2 遍历未启用）" % r.get("tried", 0)
    else:
        why = "未命中"
    print("     未找到：%s" % why, flush=True)
    print("     提示：① 该号可能确实不存在（区间是包围盒，内部有未分配号）；"
          "② 用 --meeting <会议> 在指定会议前后 ±%d 个目录内再扫；"
          "③ 用 --base <Docs 目录 URL> 手工指定。" % args.radius, flush=True)
    _print_suggest(r.get("suggest"))


def _ask_and_retry(tdoc, r, args):
    """未命中时交互式提示用户输入会议号，走 L0 重扫；支持多次尝试。"""
    _print_not_found(r, args)
    while True:
        try:
            ans = input("     请输入会议号/目录名（如 R1-126-bis 或 TSGR1_126b；回车跳过）：")
        except EOFError:
            print("     （无交互输入，跳过）", flush=True)
            return r
        ans = (ans or "").strip()
        if not ans:
            return r
        r2 = resolve_tdoc(tdoc, timeout=min(args.timeout, 25), workers=args.workers,
                          scan=False, scan_limit=args.scan_limit, base=args.base,
                          meeting=ans, radius=args.radius)
        if r2["status"] == "found":
            print("     指定会议命中：%s（%s）" % (r2["dir"], r2.get("meeting_win") or ""),
                  flush=True)
            return r2
        print("     仍未找到：%s" % (r2.get("meeting_note") or "该窗口内无此号"), flush=True)
        _print_suggest(r2.get("suggest"))
        r = r2


def run_tdocs(items, args, index):
    out_root = os.path.join(args.out, "by_tdoc")
    os.makedirs(out_root, exist_ok=True)
    for raw in items:
        tdoc = norm_tdoc_input(raw)
        if not tdoc:
            continue
        print("[TD] 定位 %s ..." % tdoc, flush=True)
        r = resolve_tdoc(tdoc, timeout=min(args.timeout, 25), workers=args.workers,
                         scan=args.scan, scan_limit=args.scan_limit,
                         force_scan=args.force_scan, base=args.base,
                         meeting=args.meeting, radius=args.radius)

        if r["status"] == "not_found" and args.ask:
            r = _ask_and_retry(tdoc, r, args)

        rec = {"kind": "tdoc", "tdoc": tdoc, "locate": r["status"], "how": r.get("how"),
               "dir": r.get("dir"), "tried": r.get("tried"),
               "rows_n": r.get("rows_n"), "n_ranged": r.get("n_ranged"),
               "scan_total": r.get("scan_total"), "note": r.get("note")}

        if r["status"] == "found":
            how = r.get("how") or ""
            if how == "meeting":
                print("     指定会议命中：%s（%s）" % (r["dir"], r.get("meeting_win") or ""),
                      flush=True)
            elif how.startswith("scan"):
                print("     兜底命中：第 %s 次探测落在 %s（候选共 %s 个目录，本次上限 %s）"
                      % (r.get("tried"), r["dir"], r.get("scan_total"), args.scan_limit),
                      flush=True)
            elif how == "range":
                print("     区间命中：%s" % r["dir"], flush=True)
            else:
                print("     手工指定：%s" % r["docs_base"], flush=True)

            if args.dry_run:
                rec["status"] = "planned"
                rec["url"] = r["docs_base"] + tdoc_zip(tdoc)
                print("     [dry-run] %s" % rec["url"], flush=True)
            else:
                res = download_one(tdoc_zip(tdoc), r["docs_base"], out_root, args.timeout)
                rec.update(res)
                msg = ("%d B" % res["size_or_err"]) if res["status"] == "ok" \
                    else str(res.get("size_or_err"))
                print("     -> %s  %s" % (res["status"], msg), flush=True)

        elif r["status"] == "not_found":
            rec["status"] = "not_found"
            if not args.ask:
                _print_not_found(r, args)

        elif r["status"] == "no_table":
            rec["status"] = "no_table"
            print("     该前缀无会议表：%s" % r.get("note"), flush=True)
            print("     提示：用 --base <Docs 目录 URL> 手工指定会议目录。", flush=True)

        else:
            rec["status"] = "unsupported"
            print("     无法识别为 TDoc 号。", flush=True)
        index.append(rec)


def run_specs(items, args, index):
    out_root = os.path.join(args.out, "specs")
    os.makedirs(out_root, exist_ok=True)
    for raw in items:
        spec = norm_spec_input(raw)
        if not spec:
            index.append({"kind": "spec", "input": raw, "status": "unsupported"})
            print("[TD] 无法识别为规范号：%s" % raw, flush=True)
            continue
        print("[TD] 规范 %s ..." % spec, flush=True)
        versions, src = fetch_spec_versions(spec)
        rec = {"kind": "spec", "spec": spec, "status": None, "source": src,
               "n_versions": len(versions)}
        if not versions:
            rec["status"] = "not_found"
            print("     未解析到任何版本（DynaReport 无该规范，且 archive 目录列举失败）。",
                  flush=True)
            index.append(rec)
            continue
        rec["latest"] = "%s (%s)" % (ver_str(versions[0]["ver"]), versions[0]["file"])
        v, err = pick_version(versions, args.spec_ver)
        if not v:
            rec["status"] = "not_found"
            rec["note"] = err
            print("     %s" % err, flush=True)
            print("     可见版本（前 8）：%s"
                  % ", ".join("%s %s" % (ver_str(x["ver"]), x["file"]) for x in versions[:8]),
                  flush=True)
            index.append(rec)
            continue
        rec["picked"], rec["picked_ver"] = v["file"], ver_str(v["ver"])
        print("     选中 %s  %s   (共 %d 个版本，来源 %s)"
              % (ver_str(v["ver"]), v["file"], len(versions), src), flush=True)
        if args.dry_run:
            rec["status"] = "planned"
            rec["url"] = v["url"]
            print("     [dry-run] %s" % v["url"], flush=True)
        else:
            res = download_one(v["file"], v["url"].rsplit("/", 1)[0] + "/", out_root, args.timeout)
            rec.update(res)
            msg = ("%d B" % res["size_or_err"]) if res["status"] == "ok" \
                else str(res.get("size_or_err"))
            print("     -> %s  %s" % (res["status"], msg), flush=True)
        index.append(rec)


def main():
    ap = argparse.ArgumentParser(
        description="3GPP 按号直下（TDoc / TS / TR）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="例：python tdoc_direct.py --direct \"R1-2607107, RP-261575, 38.211\" --out Docs")
    ap.add_argument("--get-tdoc", help="按 TDoc 号下载，支持多值(逗号/分号/顿号/空格)")
    ap.add_argument("--get-spec", help="按 TS/TR 号下载，支持多值")
    ap.add_argument("--direct", help="混合输入，自动识别 TDoc / 规范号")
    ap.add_argument("--spec-ver", help="规范版本：i40 / j50 / v18.4.0；默认取最新")
    ap.add_argument("--meeting",
                    help="[L0] 指定会议(目录名/label/序号)，在其前后 ±radius 个目录内扫描")
    ap.add_argument("--radius", type=int, default=None,
                    help="L0 前后各扫几个目录，默认 5（配置 NEIGHBOR_RADIUS）")
    ap.add_argument("--ask", action="store_true",
                    help="未命中时交互式提示输入会议号，再走 L0（CLI 用；GUI 后续接入输入框）")
    ap.add_argument("--base", help="[L3 兜底] 直给 Docs 目录 URL(以 /Docs/ 结尾)")
    ap.add_argument("--out", default=None, help="输出根目录，默认 Docs（配置 OUT）")
    ap.add_argument("--workers", type=int, default=None, help="HEAD/下载并发，默认 8（配置 WORKERS）")
    ap.add_argument("--timeout", type=int, default=None, help="下载超时秒，默认 90（配置 TIMEOUT）")
    ap.add_argument("--scan-limit", type=int, default=None,
                    help="L2 遍历会议目录的上限，默认 10（配置 SCAN_LIMIT）")
    g = ap.add_mutually_exclusive_group()
    g.add_argument("--scan", dest="scan", action="store_true", help="启用 L2 遍历（默认启用）")
    g.add_argument("--no-scan", dest="scan", action="store_false", help="关闭 L2 遍历")
    ap.set_defaults(scan=None)
    ap.add_argument("--force-scan", action="store_true",
                    help="强制走 L2 兜底（跳过 L1 区间命中，用于自测兜底链路）")
    ap.add_argument("--env", help="指定配置文件路径（默认找 tdoc_direct.env 或 .env）")
    ap.add_argument("--show-config", action="store_true", help="打印生效配置后退出")
    ap.add_argument("--no-cache", action="store_true",
                    help="本次不使用索引缓存（会议表/规范列表；仍会写入新抓到的）")
    ap.add_argument("--refresh-cache", action="store_true",
                    help="忽略已有索引缓存，重新抓取并更新")
    ap.add_argument("--cache-ttl", type=int, default=None,
                    help="索引缓存有效期秒数，默认 604800(7天)；0=永不过期（配置 CACHE_TTL）")
    ap.add_argument("--cache-dir", default=None,
                    help="索引缓存目录，默认 <下载根目录>/.tdoc_cache（配置 CACHE_DIR）")
    ap.add_argument("--clear-cache", action="store_true", help="清空索引缓存后退出")
    ap.add_argument("--dry-run", action="store_true", help="只定位、打印 URL，不下载")
    args = ap.parse_args()

    if args.env:
        load_env(args.env)
    args.workers = cfg_int("WORKERS", args.workers, 1)
    args.timeout = cfg_int("TIMEOUT", args.timeout, 1)
    args.scan_limit = cfg_int("SCAN_LIMIT", args.scan_limit, 0)
    args.radius = cfg_int("NEIGHBOR_RADIUS", args.radius, 0)
    args.out = cfg_str("OUT", args.out) or "Docs"
    if args.scan is None:
        args.scan = bool(cfg_int("SCAN", None, 0))
    # 索引缓存默认与下载产物同处：<out>/.tdoc_cache（CACHE_DIR / --cache-dir 可覆盖）
    _cache_default_dir = os.path.join(args.out, _cache.DEFAULT_SUBDIR)
    _cache.configure(
        enabled=bool(cfg_int("CACHE", None, 0)) and not args.no_cache,
        ttl=cfg_int("CACHE_TTL", args.cache_ttl, 0),
        directory=(cfg_str("CACHE_DIR", args.cache_dir) or _cache_default_dir),
        refresh=bool(args.refresh_cache))

    if args.clear_cache:
        _cache.clear()
        return

    if args.show_config:
        env = load_env()
        print("[TD] 配置文件 = %s" % (env.get("__file__") or "(无，使用默认值)"))
        print("     SCAN_LIMIT      = %s" % args.scan_limit)
        print("     NEIGHBOR_RADIUS = %s" % args.radius)
        print("     WORKERS         = %s" % args.workers)
        print("     TIMEOUT         = %s" % args.timeout)
        print("     OUT             = %s" % args.out)
        print("     SCAN            = %s" % (1 if args.scan else 0))
        ci = _cache.info()
        print("     CACHE           = %s" % (1 if ci["enabled"] else 0))
        print("     CACHE_TTL       = %s" % ci["ttl"])
        print("     CACHE_DIR       = %s" % ci["dir"])
        return

    groups = []
    if args.get_tdoc:
        groups.append(("tdoc", split_list(args.get_tdoc)))
    if args.get_spec:
        groups.append(("spec", split_list(args.get_spec)))
    if args.direct:
        # 先把 "TS 38.211" / "TR 38.821" 粘成一个词，避免被空白切散
        raw = re.sub(r"\b(TS|TR)\s+(\d{2}\.\d{3})", r"\1\2", args.direct, flags=re.I)
        t, s = [], []
        for x in split_list(raw):
            y = norm_user_input(x).strip()
            if SPEC_IN_RE.match(y) or re.fullmatch(r"\d{5}", y) or re.match(r"^(TS|TR)\d", y.upper()):
                s.append(y)
            elif TDOC_IN_RE.match(y):
                t.append(y)
            else:
                print("[TD] 忽略无法识别项：%s" % x, file=sys.stderr)
        if t:
            groups.append(("tdoc", t))
        if s:
            groups.append(("spec", s))
    if not groups:
        ap.error("至少指定 --get-tdoc / --get-spec / --direct 之一")

    if args.base:
        args.base = args.base if args.base.endswith("/") else args.base + "/"

    index = []
    for kind, items in groups:
        if kind == "tdoc":
            run_tdocs(items, args, index)
        else:
            run_specs(items, args, index)

    os.makedirs(args.out, exist_ok=True)
    idx_path = os.path.join(args.out, "direct_index.json")
    with open(idx_path, "w", encoding="utf-8") as fh:
        json.dump({"tool": "tdoc_direct", "dry_run": bool(args.dry_run),
                   "count": len(index), "items": index}, fh, ensure_ascii=False, indent=2)
    log_path = os.path.join(args.out, "_direct_log.csv")
    with open(log_path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["kind", "input", "status", "locate", "how", "dir_or_file",
                    "size_or_err", "note"])
        for r in index:
            w.writerow([r.get("kind"), r.get("tdoc") or r.get("spec") or r.get("input"),
                        r.get("status"), r.get("locate", ""), r.get("how", ""),
                        r.get("dir") or r.get("file"), r.get("size_or_err", ""),
                        r.get("note", "")])

    ok = sum(1 for r in index if r.get("status") in ("ok", "planned"))
    fail = [r for r in index if r.get("status") not in ("ok", "planned")]
    print("[TD] DONE 总=%d 成功=%d 未取得=%d" % (len(index), ok, len(fail)))
    print("[TD] 索引=%s" % idx_path)
    print("[TD] 日志=%s" % log_path)
    if fail:
        for r in fail:
            print("     - %s  [%s]  %s"
                  % (r.get("tdoc") or r.get("spec") or r.get("input"),
                     r.get("status"), r.get("note", "")))
    if index and not ok:
        sys.exit(3)


if __name__ == "__main__":
    main()
