#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""T1 — 3GPP TDoc 包抽取与按 agenda 分类工具 (rule-based, v1)

设计约定:
  - 只读取源 TDoc 包, 只在输出目录新建文件, 绝不修改/删除任何源文件或共享状态。
  - 由用户手动从 3GPP 门户下载会议 TDoc 包(zip)并解压到 --package 目录后触发本工具。

用法:
  python tdoc_extract.py --package <TDoc目录> --tdoclist <会议TDoc列表.csv> \
      --out <输出目录> [--ws <workspace>]

  --tdoclist: 会议 TDoc 列表 CSV(请从 3GPP 会议页导出为 CSV, 含 TDoc/Title/Source/Agenda/Type 等列)
  --out:      输出目录, 将创建 by_agenda/<agenda>/ 与 tdoc_index.json

输出:
  <out>/tdoc_index.json : { workspace, package, tdoc_count, agendas[], tdocs[] }
  <out>/by_agenda/<agenda>/ : 按 agenda 分类复制的 TDoc 文件
"""

import argparse
import csv
import json
import os
import re
import shutil
import sys


def _is_junk(fn):
    """macOS/系统垃圾文件(AppleDouble/DS_Store)不算正文。"""
    return fn.startswith("._") or fn in (".DS_Store", "Thumbs.db") or fn.startswith("~$")


def find_tdoc_file(package, tdoc_no):
    """在 package 目录中按 TDoc 号找文件。

    优先级:
      ① package/<tdoc_no>/  —— T0 新版按 TDoc 分子目录解压, 这里能精确定位,
         即使包内文件名不含 TDoc 号也不会匹配失败;
      ② 递归子串匹配, 整号在文件名里出现的优先。
    """
    sub = os.path.join(package, tdoc_no)
    if os.path.isdir(sub):
        cands = []
        for root, _, files in os.walk(sub):
            for f in files:
                if _is_junk(f):
                    continue
                cands.append(os.path.join(root, f))
        if cands:
            cands.sort()
            for c in cands:
                if tdoc_no.lower() in os.path.basename(c).lower():
                    return c
            return cands[0]

    pattern = re.compile(re.escape(tdoc_no), re.IGNORECASE)
    hits = []
    for root, _, files in os.walk(package):
        for f in files:
            if _is_junk(f):
                continue
            if pattern.search(f):
                hits.append(os.path.join(root, f))
    if not hits:
        return None
    for h in hits:
        if tdoc_no.lower() in os.path.basename(h).lower():
            return h
    return hits[0]


def build_field_map(fieldnames):
    """兼容不同 TDoc 列表列名, 映射到内部字段。"""
    fm = {}
    for col in fieldnames or []:
        low = (col or "").strip().lower()
        if "tdoc" in low and "list" not in low and "list" not in low:
            fm.setdefault("tdoc", col)
        if "title" in low or "subject" in low:
            fm.setdefault("title", col)
        if "source" in low or "company" in low or "origin" in low:
            fm.setdefault("source", col)
        if "agenda" in low and "sort order" not in low and "description" not in low:
            fm.setdefault("agenda", col)
        if "type" in low:
            fm.setdefault("type", col)
    return fm


def agenda_match_one(agenda, filt):
    """单个 agenda 前缀的层次化匹配: filt='9.1' 命中 9.1 / 9.1.1 / 9.1.1.2; 不会误命中 9.10。"""

    def parts(s):
        return [p.strip() for p in str(s).replace("/", "_").replace("\\", "_").split(".")
                if p.strip()]

    rp, fp = parts(agenda), parts(filt)
    if not fp:
        return True
    if len(fp) > len(rp):
        return False
    return rp[:len(fp)] == fp


def agenda_matches(agenda, filt):
    """filt 可以是单个字符串, 也可以是列表/元组/集合 —— 命中其中任意一个即算命中。"""
    if not filt:
        return True
    if isinstance(filt, (list, tuple, set, frozenset)):
        return any(agenda_match_one(agenda, f) for f in filt if f)
    return agenda_match_one(agenda, filt)


def classify_row(src, agenda, by_agenda):
    """将单个 TDoc 文件按 agenda 点号层级嵌套复制到 by_agenda 下的最细一级。

    采用累积命名: 9.1.1.1 -> by_agenda/9/9.1/9.1.1/9.1.1.1/, 与'9.1 内含有 9.1.1'的直觉一致。
    返回相对 out 的路径。
    """
    segs = [p.strip() for p in str(agenda).replace("/", "_").replace("\\", "_").split(".")
            if p.strip()]
    if not segs:
        segs = ["unknown"]
    # 累积路径: 每段 = 之前累积 + '.' + 本段
    agenda_parts = []
    for i, s in enumerate(segs):
        agenda_parts.append(s if i == 0 else agenda_parts[-1] + "." + s)
    dest_dir = os.path.join(by_agenda, *agenda_parts)
    os.makedirs(dest_dir, exist_ok=True)
    dest = os.path.join(dest_dir, os.path.basename(src))
    if not os.path.exists(dest):
        shutil.copy2(src, dest)
    # by_agenda 的父目录即 out, rel 以 out 为基准(与历史 tdoc_index.json 一致)
    return os.path.relpath(dest, os.path.dirname(by_agenda))


def classify(package, rows, out, field_map=None, agenda_filter=None):
    """对 package 中的 TDoc 文件按 agenda 嵌套分类, 输出到 <out>/by_agenda/。

    rows: list[dict](由 read_tdoclist / csv 读出的会议清单)。
    field_map: {tdoc,title,source,agenda,type} -> 列名; 缺省时按首行自动推断。
    agenda_filter: 可选 agenda 前缀(如 '9.1'), 仅保留其子树。
    返回: tdoc_index 条目列表(含 file 相对路径, 未命中文件 file='')。
    """
    by_agenda = os.path.join(out, "by_agenda")
    os.makedirs(by_agenda, exist_ok=True)
    index = []
    for r in rows:
        fm = field_map or build_field_map(r.keys())
        tdoc = (r.get(fm.get("tdoc", "")) or "").strip()
        if not tdoc:
            continue
        agenda = (r.get(fm.get("agenda", "")) or "unknown").strip() or "unknown"
        if not agenda_matches(agenda, agenda_filter):
            continue
        title = (r.get(fm.get("title", "")) or "").strip()
        source = (r.get(fm.get("source", "")) or "").strip()
        dtype = (r.get(fm.get("type", "")) or "").strip()
        src = find_tdoc_file(package, tdoc)
        rel = ""
        if src:
            rel = classify_row(src, agenda, by_agenda)
        index.append({
            "tdoc": tdoc,
            "title": title,
            "source": source,
            "agenda": agenda,
            "type": dtype,
            "file": rel,
        })
    return index


def main():
    ap = argparse.ArgumentParser(description="3GPP TDoc 包抽取与分类 (T1)")
    ap.add_argument("--package", required=True, help="TDoc 文件所在目录(已解压)")
    ap.add_argument("--tdoclist", required=True, help="会议 TDoc 列表 CSV(含 Agenda 列)")
    ap.add_argument("--out", required=True, help="输出目录(将创建 by_agenda/ 与 tdoc_index.json)")
    ap.add_argument("--ws", default=None, help="workspace 路径(仅记录, 供 skill 使用)")
    ap.add_argument("--agenda", default=None, help="只处理该 agenda 子树(可粗可细, 如 9.1 含其下全部)")
    args = ap.parse_args()

    if not os.path.isdir(args.package):
        print(f"[T1] 错误: package 目录不存在: {args.package}", file=sys.stderr)
        sys.exit(1)
    if not os.path.isfile(args.tdoclist):
        print(f"[T1] 错误: tdoclist 不存在: {args.tdoclist}", file=sys.stderr)
        sys.exit(1)

    os.makedirs(args.out, exist_ok=True)

    rows = []
    with open(args.tdoclist, encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        fm = build_field_map(reader.fieldnames)
        if "tdoc" not in fm:
            print("[T1] 警告: TDoc 列表未识别到 TDoc 列, 请检查 CSV 表头", file=sys.stderr)
        for r in reader:
            rows.append(r)

    index = classify(args.package, rows, args.out, field_map=fm, agenda_filter=args.agenda)

    out_json = os.path.join(args.out, "tdoc_index.json")
    with open(out_json, "w", encoding="utf-8") as fh:
        json.dump({
            "workspace": args.ws,
            "package": os.path.abspath(args.package),
            "tdoc_count": len(index),
            "agendas": sorted({i["agenda"] for i in index}),
            "tdocs": index,
        }, fh, ensure_ascii=False, indent=2)

    print(f"[T1] 处理完成: {len(index)} 个 TDoc, {len(set(i['agenda'] for i in index))} 个 agenda")
    print(f"[T1] 索引: {out_json}")
    print(f"[T1] 分类目录: {os.path.join(args.out, 'by_agenda')}")


if __name__ == "__main__":
    main()
