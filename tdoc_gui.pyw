#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""T0 图形界面 (最终版 v3)。

双击本文件打开窗口。零第三方依赖(纯 tkinter)。
只调用同目录的 tdoc_download.py(唯一 T0 实现)。

v3 相对 v2:
  ★ **会议自动推荐**: 选好 TDoc 清单(.xlsx)后, 从**文件名**推断 组别+会议号
     (如 TDoc_List_Meeting_RAN1#126-bis.xlsx -> RAN1 / 126-bis), 自动填好并联网核对出候选。
     文件名里没有会议信息时**不做任何猜测**, 只提示你手动填。
  ★ "保留原始 zip 与 extracted" 的文案改准: 不勾选=**已成功分类的**只留 by_agenda 一份,
     未成功分类的原件会保留下来便于排查(旧版是"有一个没成功就全都不删", 与承诺不符, 已修)。
  ★ summary 弹窗直接列出未取得明细(TDoc号 / agenda / 原因), 并指向 _failures.csv。
  ★ agenda 支持多个(逗号分隔)。

其余: 后台线程(不卡死) + 停止 + UTF-8(不乱码) + 日志落盘 + tkinter 自动切换解释器。
"""

import os
import re
import csv
import sys
import subprocess
import shutil
import threading
import queue


# ---- tkinter 缺失时自动切换到带 tkinter 的解释器(让 .pyw 双击更稳) ----
def _tkinter_ok(py):
    try:
        r = subprocess.run([py, "-c", "import tkinter"],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                           timeout=20)
        return r.returncode == 0
    except Exception:
        return False


def _relaunch():
    here = os.path.abspath(__file__)
    cands = []
    for name in ("pythonw", "python", "py"):
        pp = shutil.which(name)
        if pp and pp not in cands:
            cands.append(pp)
    for py in cands:
        if _tkinter_ok(py):
            try:
                os.execv(py, [py, here] + sys.argv[1:])
            except Exception:
                pass
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(
            0,
            u"未能找到可用的 Python(tkinter 不可用)。\n请确认已安装标准 Python(含 tkinter), "
            u"或手动运行:\n  python \"%s\"" % here,
            u"3GPP TDoc GUI 启动失败", 0x10)
    except Exception:
        pass
    sys.exit(1)


try:
    import tkinter  # noqa
except Exception:
    _relaunch()

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

TOOLS_DIR = os.path.dirname(os.path.abspath(__file__))
T0 = os.path.join(TOOLS_DIR, "tdoc_download.py")
if TOOLS_DIR not in sys.path:
    sys.path.insert(0, TOOLS_DIR)

# 优先复用 T0 里的单一来源(便于以后维护)
try:
    import tdoc_download as _T0
    HINTS = _T0.MEETING_NAMING_HINTS
    _resolve_meeting = _T0.resolve_meeting
    _recommend = _T0.recommend_from_filename
except Exception:
    HINTS = ("各组 bis 后缀写法不同: RAN1=TSGR1_126b(小写 b) / RAN2·RAN3=TSGR2_91bis(全拼 bis) / "
             "RAN4=TSGR4_94Bis(大写 Bis); SA·CT 常带城市名。\n"
             "建议点「生成候选 URL」联网核对, 或直接粘贴浏览器里打开的 Docs 目录地址。")
    _resolve_meeting = None

    def _recommend(path):
        return None, None

HINTS_UI = """【会议目录命名实例 · 实测 3GPP 真实目录, 组别决定子目录】
  RAN1  126      -> tsg_ran/WG1_RL1/TSGR1_126        RAN1  126-bis -> tsg_ran/WG1_RL1/TSGR1_126b     ★小写 b
  RAN2  91-bis   -> tsg_ran/WG2_RL2/TSGR2_91bis      RAN3  91-bis  -> tsg_ran/WG3_Iu/TSGR3_91bis      ★全拼 bis
  RAN4  94-bis   -> tsg_ran/WG4_Radio/TSGR4_94Bis    ★大写 Bis
  全会 RAN  113  -> tsg_ran/TSG_RAN/TSGR_113         SA  99        -> tsg_sa/TSG_SA/TSGS_99_Rotterdam_2023_03
【后缀各组不统一: b / bis / Bis / _e / 城市名】=> 请点「生成候选 URL」联网核对, 或直接粘贴浏览器里的 Docs 地址。"""

root = tk.Tk()
root.title("3GPP TDoc 下载/分类 (最终版 v3 · 按文件名推荐会议 · 不卡死)")
root.geometry("920x820")

# ---- 模式 ----
var_mode = tk.StringVar(value="A")   # A=下载+分类  B=仅分类
tk.Label(root, text="模式:").grid(row=0, column=0, sticky="e", padx=4, pady=4)
tk.Radiobutton(root, text="① 基于清单下载+分类", variable=var_mode, value="A").grid(row=0, column=1, sticky="w")
tk.Radiobutton(root, text="② 仅分类(已有包)", variable=var_mode, value="B").grid(row=0, column=2, sticky="w")

# ---- TDoc 清单 ----
tk.Label(root, text="TDoc 清单(xlsx/csv/zip):").grid(row=1, column=0, sticky="e", padx=4, pady=4)
e_tdoc = tk.Entry(root, width=74)
e_tdoc.grid(row=1, column=1, columnspan=2, sticky="we", padx=4)
tk.Button(root, text="浏览", command=lambda: on_browse_tdoc()).grid(row=1, column=3, padx=4)

# ---- 待分类包目录(仅分类模式) ----
tk.Label(root, text="待分类包目录(仅②):").grid(row=2, column=0, sticky="e", padx=4, pady=4)
e_pkg = tk.Entry(root, width=74)
e_pkg.grid(row=2, column=1, columnspan=2, sticky="we", padx=4)
tk.Button(root, text="浏览", command=lambda: browse_dir(e_pkg, "选择待分类文件夹")).grid(row=2, column=3, padx=4)

# ---- Docs 目录 URL(优先) ----
tk.Label(root, text="Docs 目录 URL(最高优先):").grid(row=3, column=0, sticky="e", padx=4, pady=4)
e_base = tk.Entry(root, width=74)
e_base.grid(row=3, column=1, columnspan=2, sticky="we", padx=4)
tk.Label(root, text="↑ 可直接粘贴浏览器里打开的 Docs 地址; 或留空, 用下面两行自动生成",
         fg="#555").grid(row=4, column=1, columnspan=3, sticky="w")

# ---- 自动推荐 / 生成 URL ----
tk.Label(root, text="组别 tsg:").grid(row=5, column=0, sticky="e", padx=4, pady=4)
var_tsg = tk.StringVar(value="RAN1")
tk.OptionMenu(root, var_tsg, "RAN", "RAN1", "RAN2", "RAN3", "RAN4", "SA", "CT", "GERAN").grid(row=5, column=1, sticky="w", padx=4)
tk.Label(root, text="会议号(如 126 / 126-bis):").grid(row=5, column=2, sticky="e", padx=4)
e_num = tk.Entry(root, width=18)
e_num.grid(row=5, column=3, sticky="w", padx=4)

btn_resolve = tk.Button(root, text="生成候选 URL(联网核对)", bg="#d8f0d8")
btn_resolve.grid(row=6, column=1, sticky="w", padx=4, pady=2)
btn_pick = tk.Button(root, text="使用选中候选 →填入URL")
btn_pick.grid(row=6, column=2, sticky="w", padx=4, pady=2)
tk.Label(root, text="(可双击下面列表)", fg="#555").grid(row=6, column=3, sticky="w")

_lb = tk.Listbox(root, height=4, font=("Consolas", 9))
_lb.grid(row=7, column=0, columnspan=4, sticky="we", padx=8, pady=2)
_cands = []

# ---- 输出目录 ----
tk.Label(root, text="输出目录:").grid(row=8, column=0, sticky="e", padx=4, pady=4)
e_out = tk.Entry(root, width=74)
e_out.grid(row=8, column=1, columnspan=2, sticky="we", padx=4)
tk.Button(root, text="浏览", command=lambda: browse_dir(e_out, "选择输出目录")).grid(row=8, column=3, padx=4)

# ---- agenda 过滤(支持多个) ----
tk.Label(root, text="agenda 过滤(可选):").grid(row=9, column=0, sticky="e", padx=4, pady=4)
e_agenda = tk.Entry(root, width=30)
e_agenda.grid(row=9, column=1, sticky="w", padx=4)
tk.Label(root, text="(可多个, 逗号分隔, 如 10.7,9.1; 10.7 含其下全部)", fg="#555").grid(row=9, column=2, columnspan=2, sticky="w")

# ---- 保留原始 zip ----
var_keepraw = tk.BooleanVar(value=False)
tk.Checkbutton(root, text="保留原始 zip 与 extracted（不勾选 = 已成功分类的只留 by_agenda 一份；"
                          "未成功分类的原件会保留便于排查）",
               variable=var_keepraw).grid(row=10, column=1, columnspan=3, sticky="w", padx=4, pady=4)

# ---- 运行 / 停止 / 退出 ----
btn_run = tk.Button(root, text="运行", bg="#cfe8ff")
btn_run.grid(row=11, column=1, sticky="w", padx=4, pady=6)
btn_stop = tk.Button(root, text="停止", state="disabled")
btn_stop.grid(row=11, column=2, sticky="w", padx=4, pady=6)
tk.Button(root, text="退出", command=root.destroy).grid(row=11, column=3, sticky="w", padx=4, pady=6)

# ---- 日志 ----
log = scrolledtext.ScrolledText(root, width=116, height=13)
log.grid(row=12, column=0, columnspan=4, sticky="we", padx=6, pady=6)

# ---- 命名实例(常驻显示) ----
tk.Label(root, text=HINTS_UI, justify="left", anchor="w", fg="#234",
         bg="#f4f8ff", font=("SimSun", 9)).grid(row=13, column=0, columnspan=4, sticky="we", padx=6, pady=(0, 8))

# ---- 运行态 ----
_q = queue.Queue()      # 下载子进程输出
_rq = queue.Queue()     # 联网解析候选
_state = {"proc": None, "running": False, "logfile": None, "buf": [], "last_rec": None}


def browse_file(entry, title, filetypes):
    p = filedialog.askopenfilename(title=title, filetypes=filetypes)
    if p:
        entry.delete(0, tk.END)
        entry.insert(0, p)
    return p


def browse_dir(entry, title):
    p = filedialog.askdirectory(title=title)
    if p:
        entry.delete(0, tk.END)
        entry.insert(0, p)


def on_browse_tdoc():
    p = browse_file(e_tdoc, "选择 TDoc 清单",
                    [("清单", "*.xlsx;*.xlsm;*.zip;*.csv"), ("所有文件", "*.*")])
    if p:
        auto_recommend()


def _emit(line):
    log.insert(tk.END, line)
    log.see(tk.END)
    _state["buf"].append(line)
    lf = _state["logfile"]
    if lf:
        try:
            lf.write(line)
            lf.flush()
        except Exception:
            pass


# ======================= 从文件名推荐会议 =======================
def auto_recommend(quiet=False):
    """按 TDoc 清单文件名推荐 组别+会议号; 识别不到就什么都不改。"""
    path = e_tdoc.get().strip()
    if not path:
        return
    name = os.path.basename(path)
    if _state.get("last_rec") == name:
        return
    _state["last_rec"] = name

    tsg, num = _recommend(path)
    if not tsg and not num:
        _emit(f"[界面] 文件名 '{name}' 里没有会议信息, 未做推荐 —— 请手动填组别+会议号, "
              f"或直接粘贴 Docs 目录 URL。\n")
        return
    if tsg:
        var_tsg.set(tsg)
    if num:
        e_num.delete(0, tk.END)
        e_num.insert(0, num)
    _emit(f"[界面] 由文件名推荐: 组别={tsg or '(未识别)'}  会议号={num or '(未识别)'}\n")
    if tsg and num and _resolve_meeting is not None:
        on_resolve(quiet=True)
    elif not num:
        _emit("[界面] 文件名里有组别但没识别到会议号, 请手动补上会议号。\n")


# ======================= 联网解析: 会议号 -> 候选 URL =======================
def on_resolve(quiet=False):
    num = e_num.get().strip()
    tsg = var_tsg.get().strip()
    if not num:
        if not quiet:
            messagebox.showerror("缺参数", "请填会议号, 如 126 或 126-bis")
        return
    if _resolve_meeting is None:
        if not quiet:
            messagebox.showerror("不可用", "未能加载 tdoc_download.py 的解析函数, 请直接粘贴 Docs 目录 URL。")
        return
    btn_resolve.config(state="disabled")
    _emit(f"\n[界面] 联网解析: 组别={tsg} 会议号={num} …\n")
    threading.Thread(target=_resolve_worker, args=(tsg, num, quiet), daemon=True).start()
    root.after(150, _poll_resolve)


def _resolve_worker(tsg, num, quiet):
    try:
        root_, cands = _resolve_meeting(tsg, num)
        _rq.put(("cands", (root_, cands, quiet)))
    except Exception as e:
        _rq.put(("err", (str(e), quiet)))


def _poll_resolve():
    try:
        kind, payload = _rq.get_nowait()
    except queue.Empty:
        root.after(150, _poll_resolve)
        return
    btn_resolve.config(state="normal")

    if kind == "err":
        msg, quiet = payload
        _emit(f"[界面] 解析失败: {msg}\n")
        if not quiet:
            messagebox.showerror("解析失败",
                                 f"联网解析会议目录失败:\n{msg}\n\n可改为直接粘贴 Docs 目录 URL。")
        return

    root_, cands, quiet = payload
    _emit(f"[界面] 组目录 {root_}\n")
    _lb.delete(0, tk.END)
    _cands.clear()
    if not cands:
        _emit("[界面] 未匹配到含该会议号的目录, 请检查组别/会议号\n")
        if not quiet:
            messagebox.showwarning("未找到", HINTS_UI)
        return
    for c in cands:
        tag = "DOCS" if c["has_docs"] else " -- "
        _lb.insert(tk.END, f"[{tag}] {c['folder']}    {c['docs_url']}")
        _cands.append(c)
    _lb.selection_clear(0, tk.END)
    _lb.selection_set(0)
    _emit(f"[界面] 候选 {len(cands)} 个(标签 DOCS=该会议 Docs 目录已存在/可访问)\n")
    if len(cands) == 1:
        _apply_candidate(0)
        _emit("[界面] 已自动填入 URL, 请审核后点『运行』\n")
    else:
        _emit("[界面] 有多个候选(如 126 与 126b), 请在列表里双击选中要用的那个\n")


def _apply_candidate(i):
    if 0 <= i < len(_cands):
        c = _cands[i]
        e_base.delete(0, tk.END)
        e_base.insert(0, c["docs_url"])
        _emit(f"[界面] 已填入 URL(请核对): {c['docs_url']}\n")


def on_pick():
    sel = _lb.curselection()
    if not sel:
        messagebox.showerror("未选择", "请先在候选列表里选中一行")
        return
    _apply_candidate(sel[0])


def on_lb_double(event=None):
    on_pick()


# ======================= 运行 T0 =======================
def _worker(cmd):
    try:
        env = dict(os.environ)
        env["PYTHONIOENCODING"] = "utf-8"   # 让子进程按 UTF-8 输出, 避免中文乱码
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                text=True, encoding="utf-8", errors="replace", bufsize=1, env=env)
        _state["proc"] = proc
        for line in proc.stdout:
            _q.put(("line", line))
        proc.wait()
        _q.put(("done", proc.returncode))
    except Exception as e:
        _q.put(("err", str(e)))


def _poll():
    try:
        while True:
            kind, payload = _q.get_nowait()
            if kind == "line":
                _emit(payload)
            elif kind == "done":
                _finish(payload)
                return
            elif kind == "err":
                _emit("\n运行失败: " + str(payload) + "\n")
                _finish(-1, err=str(payload))
                return
    except queue.Empty:
        pass
    if _state["running"]:
        root.after(80, _poll)


def _read_failures(out_dir, limit=12):
    """从 _failures.csv 读出未取得明细, 返回可显示的短行。"""
    fp = os.path.join(out_dir, "_failures.csv")
    rows = []
    try:
        with open(fp, encoding="utf-8-sig", newline="") as f:
            for i, r in enumerate(csv.DictReader(f)):
                if i >= limit:
                    break
                rows.append("  {tdoc:<14} agenda={ag:<10} {rs}".format(
                    tdoc=(r.get("tdoc") or "")[:14],
                    ag=(r.get("agenda(s)") or "-")[:10],
                    rs=(r.get("reason") or "")))
    except Exception:
        pass
    return rows, fp


def _finish(rc, err=None):
    _state["running"] = False
    _state["proc"] = None
    btn_run.config(state="normal")
    btn_stop.config(state="disabled")
    text = "".join(_state["buf"])
    log.insert(tk.END, f"\n=== 退出码: {rc} ===\n")
    out_dir = _state.get("out", "")
    log_path = _state.get("log_path", "")
    lf = _state["logfile"]
    if lf:
        try:
            lf.write(f"\n=== 退出码: {rc} ===\n")
            lf.close()
        except Exception:
            pass
        _state["logfile"] = None
    if err:
        messagebox.showerror("错误", err)
        return

    # ---- 精确判定: 退出码 + T0 的 DONE 统计(不要用"文本里有没有'失败'两字") ----
    m_total = re.search(r"总=(\d+)", text)
    m_valid = re.search(r"有效=(\d+)", text)
    m_fail = re.search(r"仍失败=(\d+)", text)
    m_reason = re.search(r"失败原因统计:\s*(.+)", text)
    total = int(m_total.group(1)) if m_total else None
    valid = int(m_valid.group(1)) if m_valid else None
    still = int(m_fail.group(1)) if m_fail else 0
    reason = m_reason.group(1).strip() if m_reason else ""

    if rc == 0:
        head = "✅ 运行完成(退出码 0)"
        if total is not None and valid is not None:
            head += f"\n下载: 有效 {valid}/{total}"
        if still:
            head += f"\n未取得: {still} 个" + (f"  ({reason})" if reason else "")
            rows, fp = _read_failures(out_dir)
            if rows:
                head += "\n\n未取得明细(TDoc / agenda / 原因):\n" + "\n".join(rows)
            head += ("\n\n→ 多为 403/404: 该文件在 3GPP 侧不公开或不存在, **不影响其余文件**, "
                     "不是脚本出错。")
            if rows:
                head += f"\n完整明细(含标题/来源/URL): {fp}"
        else:
            head += "\n全部文件均已取得。"
        head += f"\n\n输出目录:\n{os.path.abspath(out_dir)}\n日志: {log_path}"
        messagebox.showinfo("完成", head)
    else:
        messagebox.showwarning(
            "运行未完成",
            f"脚本退出码 {rc}(非 0, 属真实错误)。\n日志:\n{log_path}\n\n"
            "常见原因:\n"
            "• Docs 目录 URL 不对 —— 点「生成候选 URL(联网核对)」可避免;\n"
            "• 网络不可达 / 代理问题;\n"
            "• 整个会议文档尚未公开(预检会在几秒内直接停止)。\n\n"
            "可把 tdoc_gui_run.log 的内容贴给我, 我帮你精确定位。")


def on_run():
    if _state["running"]:
        return
    mode = var_mode.get()
    tdoclist = e_tdoc.get().strip()
    package = e_pkg.get().strip()
    base_url = e_base.get().strip()
    out = e_out.get().strip()
    agenda = e_agenda.get().strip()
    keep_raw = var_keepraw.get()
    tsg = var_tsg.get().strip()

    if not os.path.isfile(T0):
        messagebox.showerror("缺文件", f"找不到 T0 脚本:\n{T0}")
        return
    if not out:
        messagebox.showerror("缺参数", "请填写输出目录")
        return

    cmd = [sys.executable, T0]
    if mode == "B":
        if not package or not tdoclist:
            messagebox.showerror("缺参数", "仅分类模式需要: 待分类包目录 + TDoc 清单")
            return
        cmd += ["--classify", "--package", package, "--from-tdoclist", tdoclist]
    else:
        if not tdoclist:
            messagebox.showerror("缺参数", "请填写 TDoc 清单(xlsx/csv/zip)")
            return
        if base_url:
            cmd += ["--from-tdoclist", tdoclist, "--base", base_url]
        else:
            messagebox.showerror(
                "缺参数",
                "请先点「生成候选 URL(联网核对)」并选中候选(会自动填入 URL), "
                "或直接粘贴 Docs 目录 URL。")
            return
    cmd += ["--out", out]
    if agenda:
        cmd += ["--agenda", agenda]
    if keep_raw:
        cmd += ["--keep-raw"]

    log.delete(1.0, tk.END)
    _state["buf"] = []
    _state["out"] = out
    os.makedirs(out, exist_ok=True)
    log_path = os.path.join(out, "tdoc_gui_run.log")
    _state["log_path"] = log_path
    try:
        _state["logfile"] = open(log_path, "w", encoding="utf-8")
    except Exception:
        _state["logfile"] = None

    _emit(">>> " + " ".join(cmd) + "\n")
    _state["running"] = True
    btn_run.config(state="disabled")
    btn_stop.config(state="normal")
    threading.Thread(target=_worker, args=(cmd,), daemon=True).start()
    root.after(80, _poll)


def on_stop():
    proc = _state["proc"]
    if proc and proc.poll() is None:
        try:
            proc.terminate()
        except Exception:
            pass
        _emit("\n[界面] 已请求停止 …\n")


btn_run.config(command=on_run)
btn_stop.config(command=on_stop)
btn_resolve.config(command=on_resolve)
btn_pick.config(command=on_pick)
_lb.bind("<Double-Button-1>", on_lb_double)
e_tdoc.bind("<FocusOut>", lambda ev: auto_recommend())

root.mainloop()
