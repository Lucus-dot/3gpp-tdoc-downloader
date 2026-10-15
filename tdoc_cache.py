#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""索引缓存 —— 3GPP「会议表 / 规范版本列表」的本地磁盘缓存（纯标准库）。

背景：tdoc_direct.py 定位时要抓几类**索引页**（单页可达 400KB）：
  - 会议表    DynaReport/Meetings-<code>.htm        （TDoc 号区间 → 会议目录）
  - 规范列表  DynaReport/<digits>.htm               （TS/TR 号 → 各版本 zip）
  - 目录列举  ftp/Specs/archive/<s>_series/<spec>/  （上一张表的兜底）
这些索引在同一天内**重复运行时内容不变**，因此落盘缓存，换进程 / 重启 GUI 都不必重抓。

接入方式（在 tdoc_direct.py 里一行包住 _get_text 即可）:
    import tdoc_cache as _cache
    _get_text = _cache.wrap(_get_text)

配置（优先级 CLI > 配置文件 > 默认）:
    CACHE=1             # 1 启用（默认）/ 0 关闭
    CACHE_TTL=604800    # 有效期秒，默认 7 天；0 或负数 = 永不过期
    CACHE_DIR=          # 留空 = <下载根目录>/.tdoc_cache（即 --out 目录下，与下载产物同处）

缓存目录解析（按优先级，逐个尝试创建，取第一个可写的）:
    1. CACHE_DIR / --cache-dir 显式指定
    2. <下载根目录>/.tdoc_cache        ← 默认（由 tdoc_direct.py 传入）
    3. <脚本目录>/.tdoc_cache          ← 兜底（下载目录不可写时）
    4. <系统临时目录>/tdoc_direct_cache ← 最后兜底

边界（重要）:
  - **只缓存索引页**。真正的 zip 下载（tdoc_download.fetch）与 HEAD 存在性校验
    （tdoc_direct._head_ok）**一律不过缓存** —— 后者是"这个文件到底存不存在"的判据，必须实时。
  - 写入是原子的（先写 .tmp 再 os.replace），多进程并发不会读到半截文件。
  - 内容过短（<64 字节，多半是错误页）不写缓存。
"""

import os
import re
import tempfile
import time

DEFAULTS = {"CACHE": 1, "CACHE_TTL": 604800}

CFG = {"enabled": True, "ttl": 604800, "dir": None, "refresh": False, "quiet": False}
_HITS = {"n": 0, "max_age": None}

# 索引页 URL → 缓存键（只认这三种；其余 URL 一律直连，不缓存）
_M_RE = re.compile(r"/DynaReport/Meetings-([A-Za-z]{1,2}\d?)\.htm", re.I)
_S_RE = re.compile(r"/DynaReport/(\d{5})\.htm", re.I)
_L_RE = re.compile(r"/Specs/archive/(\d+)_series/(\d+\.\d+)/?$", re.I)

MIN_CACHE_CHARS = 64          # 太短（多半是错误页）不写缓存
DEFAULT_SUBDIR = ".tdoc_cache"


# ===========================================================================
# 配置 / 路径
# ===========================================================================
def configure(enabled=True, ttl=604800, directory=None, refresh=False, quiet=False):
    global CFG
    CFG["enabled"] = bool(enabled)
    try:
        CFG["ttl"] = int(ttl)
    except (TypeError, ValueError):
        CFG["ttl"] = DEFAULTS["CACHE_TTL"]
    # 指定目录统一转绝对路径，避免工作目录变动导致落点漂移
    CFG["dir"] = os.path.abspath(directory) if directory else None
    CFG["refresh"] = bool(refresh)
    CFG["quiet"] = bool(quiet)
    return CFG


def _candidate_dirs():
    """缓存目录候选（按优先级去重）: 指定目录 → 脚本目录 → 系统临时目录。"""
    cands = []
    if CFG["dir"]:
        cands.append(CFG["dir"])
    cands.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), DEFAULT_SUBDIR))
    cands.append(os.path.join(tempfile.gettempdir(), "tdoc_direct_cache"))
    seen, uniq = set(), []
    for d in cands:
        k = os.path.normcase(os.path.abspath(d))
        if k not in seen:
            seen.add(k)
            uniq.append(d)
    return uniq


def cache_dir(create=True):
    """解析当前缓存目录。
    create=False → 只给首选路径（报告用，不落盘）。
    create=True  → 按候选链逐个尝试创建，返回第一个可写目录；全失败返回 None。
    """
    cands = _candidate_dirs()
    if not create:
        return cands[0]
    for d in cands:
        try:
            os.makedirs(d, exist_ok=True)
            CFG["dir"] = d
            return d
        except OSError:
            continue
    return None


def _path(key):
    d = cache_dir(create=True)
    return os.path.join(d, key + ".txt") if d else None


def key_for(url):
    """索引页 URL → 缓存键；非索引页返回 None（不缓存）。"""
    m = _M_RE.search(url)
    if m:
        return "meetings_" + m.group(1).upper()
    m = _S_RE.search(url)
    if m:
        return "spec_" + m.group(1)
    m = _L_RE.search(url)
    if m:
        return "speclist_%s_%s" % (m.group(1), m.group(2).replace(".", "_"))
    return None


def fmt_age(sec):
    if sec is None:
        return "未知"
    s = int(sec)
    if s < 60:
        return "%d 秒" % s
    if s < 3600:
        return "%d 分钟" % (s // 60)
    if s < 86400:
        return "%.1f 小时" % (s / 3600.0)
    return "%.1f 天" % (s / 86400.0)


# ===========================================================================
# 读 / 写 / 清
# ===========================================================================
def read(key):
    """命中返回 (text, age_sec)；未命中/过期/禁用返回 (None, None)。"""
    if not CFG["enabled"] or CFG["refresh"]:
        return None, None
    p = _path(key)
    if not p or not os.path.isfile(p):
        return None, None
    try:
        age = max(0, int(time.time() - os.path.getmtime(p)))
    except OSError:
        return None, None
    ttl = CFG["ttl"]
    if isinstance(ttl, int) and ttl > 0 and age > ttl:
        return None, None
    try:
        with open(p, "r", encoding="utf-8", errors="replace") as fh:
            txt = fh.read()
    except OSError:
        return None, None
    return (txt, age) if txt else (None, None)


def write(key, text):
    if not CFG["enabled"] or not text or len(text) < MIN_CACHE_CHARS:
        return False
    p = _path(key)
    if not p:
        return False
    try:
        os.makedirs(os.path.dirname(p), exist_ok=True)
        tmp = "%s.tmp%d" % (p, os.getpid())
        with open(tmp, "w", encoding="utf-8") as fh:
            fh.write(text)
        os.replace(tmp, p)                # 原子替换，避免读到半截文件
        return True
    except OSError:
        return False


def wrap(orig):
    """把 _get_text 包一层磁盘缓存。签名与 _get_text(url, timeout=45) 一致。"""
    def cached(url, timeout=45):
        key = key_for(url)
        if key is None or not CFG["enabled"]:
            return orig(url, timeout=timeout)
        txt, age = read(key)
        if txt is not None:
            _HITS["n"] += 1
            if age is not None and (age > (_HITS["max_age"] or 0)):
                _HITS["max_age"] = age
            if not CFG["quiet"]:
                print("[TD] 索引缓存命中 %s（%s前抓取；要最新请加 --refresh-cache）"
                      % (key, fmt_age(age)), flush=True)
            return txt
        txt = orig(url, timeout=timeout)
        write(key, txt)
        return txt
    return cached


def clear():
    """清理所有候选目录里本模块写下的缓存文件（meetings_/spec_/speclist_ 前缀的 .txt 与 .tmp）。"""
    dirs = [d for d in _candidate_dirs() if os.path.isdir(d)]
    if not dirs:
        print("[TD] 缓存目录不存在，无需清理。")
        return 0
    n = 0
    for d in dirs:
        for fn in os.listdir(d):
            if not re.match(r"^(meetings_|spec_|speclist_)", fn):
                continue
            if not (fn.endswith(".txt") or ".tmp" in fn):
                continue
            try:
                os.remove(os.path.join(d, fn))
                n += 1
            except OSError:
                pass
    print("[TD] 已清理索引缓存 %d 个文件（目录：%s）" % (n, " ; ".join(dirs)))
    return n


def info():
    return {"enabled": CFG["enabled"], "ttl": CFG["ttl"],
            "dir": (cache_dir(create=False) if CFG["enabled"] else "(已关闭)"),
            "refresh": CFG["refresh"], "hits": _HITS["n"]}


def hit_info():
    """(命中次数, 命中缓存的最大年龄秒) —— 供调用方提示"用的是旧索引"。"""
    return _HITS["n"], _HITS["max_age"]
