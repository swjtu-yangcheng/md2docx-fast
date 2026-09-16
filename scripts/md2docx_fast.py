# -*- coding: utf-8 -*-
"""
md2docx_fast.py —— Markdown → 排版 Word 的**单进程批处理**流水线。

设计要点（相对旧流程的优化）：
1. 单进程 import edsdk，直接走 JSON-RPC：省掉「每次工具调用 = 一次 python 冷启动(~0.6s) + 一次 Agent 往返」。
2. 只解析一次文档结构（doc_resolve_document_structure），所有插入点用已知索引计算，
   不再每插一次就 doc_find 重新定位（旧流程插入后索引漂移，被迫反复查找）。
3. 同一索引处多次插入时按「逆序插入」构造内容顺序，索引全程有效、无需重算。
4. 标题/表格按级别与表批量下发 ranges/cells，而非逐段逐表调用。
5. 默认跳过 doc_to_image（云端转换不稳定且最耗时）；需要时用 --render 显式开启。
"""
import argparse
import contextlib
import importlib.util
import io
import json
import os
import re
import subprocess
import tempfile
import sys
import time

def _detect(suffix):
    """按常见安装位置定位 editor_sdk 的 skill 目录 / 二进制（可用环境变量覆盖）。"""
    user = os.environ.get("USERNAME", "")
    cands = [
        os.path.join(os.path.expanduser("~"), "AppData", "Local", "Programs",
                     "WorkBuddy", "resources", "app.asar.unpacked", *suffix.split("\\")),
        os.path.join("D:", os.sep, "Users", user, "AppData", "Local", "Programs",
                     "WorkBuddy", "resources", "app.asar.unpacked", *suffix.split("\\")),
        os.path.join("C:", os.sep, "Program Files", "WorkBuddy", "resources",
                     "app.asar.unpacked", *suffix.split("\\")),
    ]
    for c in cands:
        if os.path.exists(c):
            return c
    return cands[0]


SDK_DIR = os.environ.get("EDSDK_DIR", "") or _detect(
    r"resources\plugins\workbuddy-builtin\skills\tencent-local-office-edit")
EDITOR_SDK_EXE = os.environ.get("EDITOR_SDK_EXE", "") or _detect(
    r"node_modules\@tencent\tencent-docs-ai-engine\bin\win32-x64\editor_sdk.exe")
SERVICE_PORT = os.environ.get("EDITOR_SDK_PORT", "39099")

# ---------- 加载 edsdk 模块（不启子进程） ----------
_spec = importlib.util.spec_from_file_location("edsdk", os.path.join(SDK_DIR, "edsdk.py"))
edsdk = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(edsdk)

_TIMINGS = []
_NOT_OPEN_HINTS = ("is not open", "not opened", "document is not")


def _raw(tool, arguments):
    """发一次 RPC；失败时把错误文本抛成 RuntimeError，避免 sys.exit 打断流水线。"""
    buf_out, buf_err = io.StringIO(), io.StringIO()
    try:
        with contextlib.redirect_stdout(buf_out), contextlib.redirect_stderr(buf_err):
            res = edsdk._rpc("tools/call", {"name": tool, "arguments": arguments})
    except SystemExit as e:
        raise RuntimeError((buf_out.getvalue() or buf_err.getvalue() or f"exit {e.code}").strip())
    texts = [c.get("text", "") for c in (res.get("content") or []) if isinstance(c, dict)]
    return "\n".join(texts) if texts else json.dumps(res, ensure_ascii=False)


def service_alive(timeout=2.0):
    """探测 editor_sdk 是否活着（tools/list 探针）。"""
    try:
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            edsdk._request(f"http://127.0.0.1:{SERVICE_PORT}/mcp", "tools/list", timeout=timeout)
        return True
    except Exception:  # noqa: BLE001
        return False


def ensure_service(timeout=15.0):
    """服务挂了就地拉起：editor_sdk.exe --port <port>（冷启动约 0.7s）。"""
    if service_alive():
        return 0.0
    t0 = time.perf_counter()
    log_path = os.path.join(tempfile.gettempdir(), "editor_sdk.log")
    with open(log_path, "a", encoding="utf-8", errors="replace") as log:
        subprocess.Popen([EDITOR_SDK_EXE, "--port", SERVICE_PORT],
                         cwd=os.path.dirname(EDITOR_SDK_EXE), stdout=log,
                         stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL,
                         creationflags=0x00000008 | 0x00000200)  # DETACHED|NEW_GROUP
    t_end = time.perf_counter() + timeout
    while time.perf_counter() < t_end:
        time.sleep(0.25)
        if service_alive():
            return time.perf_counter() - t0
    return -1.0


def wait_ready(fid, timeout=15.0, interval=0.1):
    """轮询直到文档真正 open（create_doc 返回后存在 ~0.5s 竞态窗口）。"""
    t0 = time.perf_counter()
    while time.perf_counter() - t0 < timeout:
        try:
            _raw("doc_get_outline", {"file_id": fid})
            return True
        except RuntimeError as e:
            if not any(h in str(e) for h in _NOT_OPEN_HINTS):
                return False
            time.sleep(interval)
    return False


def call(tool, arguments, _retry=True):
    """一次 JSON-RPC 调用；遇到 'document is not open' 自动等待就绪后重试。"""
    t0 = time.perf_counter()
    try:
        out = _raw(tool, arguments)
    except RuntimeError as e:
        msg = str(e)
        if _retry and any(h in msg for h in _NOT_OPEN_HINTS):
            fid = arguments.get("file_id")
            if fid and wait_ready(fid):
                out = _raw(tool, arguments)          # 就绪后重试一次
            else:
                raise
        else:
            raise
    _TIMINGS.append((tool, time.perf_counter() - t0))
    return out


def step(name, tool, arguments):
    """带步骤名计时 + 失败不中断的调用包装。"""
    try:
        out = call(tool, arguments)
    except Exception as e:  # noqa: BLE001
        print(f"  [FAIL] {name}: {str(e)[:160]}")
        return None
    snippet = out.strip().replace("\n", " ")
    print(f"  [ok] {name} ({len(out)}B) {snippet[:90]}")
    return out


def grab_file_id(text):
    """从 create_doc 返回中抽取 file_id（兼容纯文本 / JSON 两种返回）。"""
    if not text:
        return None
    s = text.strip()
    try:
        obj = json.loads(s)
        for k in ("file_id", "fileId", "id"):
            if isinstance(obj, dict) and obj.get(k):
                return obj[k]
    except Exception:  # noqa: BLE001
        pass
    m = re.search(r"(new_doc_[0-9]+_[0-9a-z]+|[0-9a-f-]{8,})", s)
    return m.group(1) if m else None


HEADING_STYLES = {
    1: {"para": {"spacing_before": 12, "spacing_after": 18, "jc": "center"},
        "text": {"font_family": "SimHei", "font_size": 20, "bold": True, "color": "000000"}},
    2: {"para": {"spacing_before": 18, "spacing_after": 9},
        "text": {"font_family": "SimHei", "font_size": 16, "bold": True, "color": "1F4E79"}},
    3: {"para": {"spacing_before": 12, "spacing_after": 6},
        "text": {"font_family": "SimHei", "font_size": 14, "bold": True, "color": "2E74B5"}},
    4: {"para": {"spacing_before": 9, "spacing_after": 6},
        "text": {"font_family": "SimHei", "font_size": 12.5, "bold": True, "color": "404040"}},
}

TBL_BORDERS = {
    "top": {"style": "single", "size": 6, "color": "8EAADB"},
    "bottom": {"style": "single", "size": 6, "color": "8EAADB"},
    "left": {"style": "single", "size": 4, "color": "BFBFBF"},
    "right": {"style": "single", "size": 4, "color": "BFBFBF"},
    "inside_h": {"style": "single", "size": 4, "color": "BFBFBF"},
    "inside_v": {"style": "single", "size": 4, "color": "BFBFBF"},
}


def parse_structure(raw):
    """doc_resolve_document_structure 返回可能是 JSON 或嵌套结构，尽力提取 nodes。"""
    try:
        data = json.loads(raw)
    except Exception:  # noqa: BLE001
        return []
    if isinstance(data, list):
        return data
    for key in ("nodes", "items", "structure", "elements"):
        v = data.get(key)
        if isinstance(v, list):
            return v
    return []


def first_range(raw):
    """doc_find 返回形如 {"locations":[{begin,end,...}]}（也可能是 ranges / 裸列表）。"""
    if not raw:
        return None
    try:
        obj = json.loads(raw)
    except Exception:  # noqa: BLE001
        return None
    cand = None
    if isinstance(obj, list):
        cand = obj
    elif isinstance(obj, dict):
        for key in ("locations", "ranges", "matches", "results", "items"):
            if isinstance(obj.get(key), list):
                cand = obj[key]
                break
        if cand is None and "begin" in obj:
            cand = [obj]
    for r in cand or []:
        if isinstance(r, dict) and "begin" in r and "end" in r:
            return r
    return None


def main():
    ap = argparse.ArgumentParser(description="Markdown → 排版 Word（单进程快速流水线）")
    ap.add_argument("md_path", help="源 Markdown 文件（本地路径）")
    ap.add_argument("out_path", help="输出 .docx 路径")
    ap.add_argument("--title", default="", help="页眉文字（缺省取首个一级标题）")
    ap.add_argument("--toc-anchor", default="摘要", help="目录插入锚点：该标题之前插入目录页")
    ap.add_argument("--toc-level", type=int, default=3, help="目录层级，默认 3")
    ap.add_argument("--render", action="store_true", help="额外渲染逐页图片（云端服务，慢且不稳定）")
    args = ap.parse_args()

    t_start = time.perf_counter()
    t_svc = ensure_service()
    print(f"[0/9] editor_sdk 服务就绪 ({'已存活' if t_svc == 0 else f'重启耗时 {t_svc:.2f}s'})")
    md_uri = "file://" + os.path.abspath(args.md_path).replace("\\", "/")

    print("[1/9] 新建文档")
    fid = grab_file_id(step("create_doc", "create_doc", {}))
    if not fid:
        print("!! 无法获取 file_id，终止")
        return 1
    t_r = time.perf_counter()
    wait_ready(fid)
    print(f"  file_id = {fid}  (就绪等待 {time.perf_counter()-t_r:.2f}s)")

    print("[2/9] 写入 Markdown 全文")
    step("doc_insert_markdown", "doc_insert_markdown",
         {"file_id": fid, "idx": 0, "markdown": md_uri})

    print("[3/9] 解析文档结构（一次性，后续全部基于索引计算）")
    nodes = parse_structure(step("doc_resolve_document_structure", "doc_resolve_document_structure",
                                 {"file_id": fid}))
    headings = [n for n in nodes if n.get("type") == "Heading"]
    print(f"  标题 {len(headings)} 个")

    print("[4/9] 批量应用标题样式（按级别一次性下发 ranges）")
    by_level = {}
    for h in headings:
        try:
            lvl = int(h.get("heading_level"))
        except (TypeError, ValueError):
            continue
        by_level.setdefault(lvl, []).append(
            {"begin": int(h["start_index"]), "end": int(h["end_index"]) - 1})
    for lvl in sorted(k for k in by_level if k in HEADING_STYLES):
        st = HEADING_STYLES[lvl]
        ranges = by_level[lvl]
        step(f"H{lvl} 段落({len(ranges)})", "doc_modify_paragraph",
             {"file_id": fid, "ranges": ranges, **st["para"]})
        step(f"H{lvl} 文字({len(ranges)})", "doc_update_text_property",
             {"file_id": fid, "ranges": ranges, **st["text"]})

    print("[5/9] 插入目录页（逆序插入，索引不漂移）")
    # 结构里的标题文本字段是 text_preview（可能被截断），匹配需双向包含
    def _hit(h):
        t = str(h.get("text_preview") or h.get("text") or "")
        return args.toc_anchor in t or t in args.toc_anchor

    anchor = next((h for h in headings if _hit(h)), None)
    if anchor is None:
        # 兜底：用 doc_find 精确查找锚点（返回字段是 locations）
        found = step("兜底定位锚点", "doc_find", {"file_id": fid, "text": args.toc_anchor})
        r0 = first_range(found)
        if r0:
            anchor = {"start_index": int(r0["begin"])}
    if anchor:
        a = int(anchor["start_index"])
        label = "目  录"
        # ★ 顺序至关重要：目录域必须落在「标签+空段」之后。
        # 若在刚插入分页符的同一索引处插入目录域，editor_sdk 会崩溃（连接被强制关闭 → 服务退出）。
        step("目录标签", "doc_insert_text", {"file_id": fid, "idx": a, "text": label})
        step("空段", "doc_insert_paragraph", {"file_id": fid, "idx": a + len(label), "level": 0})
        step("目录域", "doc_insert_toc",
             {"file_id": fid, "idx": a + len(label) + 1, "max_level": args.toc_level})
        # 摘要前分页：锚点已后移，查一次定位
        r1 = first_range(step("定位摘要", "doc_find", {"file_id": fid, "text": args.toc_anchor}))
        if r1:
            step("分页(摘要前)", "doc_insert_page_break",
                 {"file_id": fid, "idx": int(r1["begin"])})
        step("分页(目录前)", "doc_insert_page_break", {"file_id": fid, "idx": a})
        # 标签样式（查一次定位，比推算索引更稳）
        rl = first_range(step("定位目录标签", "doc_find", {"file_id": fid, "text": "目"}))
        if rl:
            ranges = [{"begin": int(rl["begin"]), "end": int(rl["end"])}]
            step("标签段落居中", "doc_modify_paragraph",
                 {"file_id": fid, "ranges": ranges, "jc": "center",
                  "spacing_before": 12, "spacing_after": 12})
            step("标签文字", "doc_update_text_property",
                 {"file_id": fid, "ranges": ranges, "bold": True,
                  "font_size": 16, "font_family": "SimHei"})
    else:
        print(f"  (未找到锚点标题 {args.toc_anchor!r}，跳过目录)")

    print("[6/9] 文档级样式（字体/行距/页边距）")
    step("doc_set_document_style", "doc_set_document_style", {"file_id": fid, **{
        "default_text_style": {"font_family": "SimSun", "font_size": 12},
        "default_paragraph_style": {"line_spacing": 1.5, "line_spacing_rule": 1},
        "page_style": {"top_margin": 72, "bottom_margin": 72,
                       "left_margin": 90, "right_margin": 90}}})

    print("[7/9] 表格美化（逐表批量）")
    raw = step("doc_list_tables", "doc_list_tables", {"file_id": fid})
    tables = []
    try:
        obj = json.loads(raw)
        if isinstance(obj, list):
            tables = obj
        elif isinstance(obj, dict):
            tables = obj.get("tables") or obj.get("items") or []
    except Exception:  # noqa: BLE001
        pass
    for i, tb in enumerate(tables, 1):
        tid = tb.get("table_id") or tb.get("id")
        cols = int(tb.get("col_count") or tb.get("column_count") or 0)
        if not tid:
            continue
        step(f"表{i} 样式", "doc_set_table_properties", {"file_id": fid, "table_id": tid, **{
            "alignment": "center", "cell_v_align": "center",
            "cell_margin": {"top": 40, "bottom": 40, "left": 80, "right": 80},
            "borders": TBL_BORDERS,
            "cell_fills": [{"condition": "first_row", "fill": "DCE6F1"}]}})
        if cols:
            step(f"表{i} 表头加粗", "doc_set_table_cells", {"file_id": fid, "table_id": tid, "cells": [
                {"row": 1, "col": c, "text_format": {"bold": True}} for c in range(1, cols + 1)]})

    print("[8/9] 页眉 + 页码")
    title = args.title or ""
    if not title:
        h1 = next((h for h in headings if int(h.get("heading_level", 0)) == 1), None)
        title = str(h1.get("text_preview") or h1.get("text") or "").strip() if h1 else \
            os.path.splitext(os.path.basename(args.md_path))[0]
    step("doc_insert_header", "doc_insert_header",
         {"file_id": fid, "text": title, "horizontal_align": "center"})
    step("doc_set_page_number", "doc_set_page_number", {"file_id": fid})

    print("[9/9] 保存")
    step("save_file", "save_file", {"file_id": fid, "file_path": os.path.abspath(args.out_path)})

    if args.render:
        print("[extra] 渲染预览图（云端转换，可能失败）")
        ascii_copy = os.path.join(os.path.dirname(os.path.abspath(args.out_path)), "_render_tmp.docx")
        try:
            import shutil
            shutil.copy(os.path.abspath(args.out_path), ascii_copy)
            step("doc_to_image", "doc_to_image", {"file_path": ascii_copy,
                                                  "output_dir": os.path.dirname(ascii_copy)})
        except Exception as e:  # noqa: BLE001
            print(f"  (渲染失败，忽略: {e})")

    total = time.perf_counter() - t_start
    print("\n===== 耗时明细 =====")
    for name, dt in _TIMINGS:
        print(f"  {dt*1000:8.1f} ms  {name}")
    print(f"  {'-'*30}\n  调用次数: {len(_TIMINGS)}   总耗时: {total:.2f} s")
    return 0


if __name__ == "__main__":
    sys.exit(main())
