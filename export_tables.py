# -*- coding: utf-8 -*-
"""
部分导出配置表（支持 1~N 张）—— 核心库，供 app.py 复用，也可命令行运行。

工作目录(work_dir) = 存放 Index.xlsx / *.xlsx / out/ 的配置表目录。
命令行用法(在配置表目录下运行，或用 --dir 指定):
  python export_tables.py <表名...> [选项]
  python export_tables.py --dir <工作目录> <表名...> [选项]
  python export_tables.py --list            # 列出全部表名
  python export_tables.py --list hero       # 列出含 hero 的表
"""

import sys
import os
import subprocess
from openpyxl import load_workbook

# 工具目录(tabtoy.exe 所在)；工作目录默认取当前目录，app.py 会覆盖
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TABTOY = os.path.join(SCRIPT_DIR, "tabtoy.exe")
WORK_DIR = os.getcwd()


def index_file():
    return os.path.join(WORK_DIR, "Index.xlsx")


def temp_index():
    return os.path.join(WORK_DIR, "_temp_index.xlsx")


def out_dir():
    return os.path.join(WORK_DIR, "out")


def load_index():
    """返回 (header, type_rows, data_rows)。"""
    wb = load_workbook(index_file(), data_only=True)
    ws = wb["index"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    header = rows[0]
    type_rows = [r for r in rows[1:] if r[0] and not r[1]]   # A列非空 且 B列空
    data_rows = [r for r in rows[1:] if r[1]]                 # B列非空 => 数据表
    return header, type_rows, data_rows


def resolve_tables(data_rows, queries):
    """按表名匹配，支持精确(忽略大小写) -> 前缀/包含匹配。返回选中的行列表。"""
    by_name = {str(r[1]).strip().lower(): r for r in data_rows}
    selected = []
    for q in queries:
        ql = q.strip().lower()
        if not ql:
            continue
        if ql in by_name:
            selected.append(by_name[ql])
            continue
        # 前缀匹配
        hits = [r for r in data_rows if str(r[1]).strip().lower().startswith(ql)]
        if len(hits) == 1:
            selected.append(hits[0])
            continue
        if len(hits) > 1:
            names = ", ".join(str(r[1]) for r in hits[:15])
            print(f"[提示] '{q}' 命中多张表({len(hits)}张): {names}")
            return None
        # 包含匹配
        hits = [r for r in data_rows if ql in str(r[1]).strip().lower()]
        if len(hits) == 1:
            selected.append(hits[0])
            continue
        if len(hits) > 1:
            names = ", ".join(str(r[1]) for r in hits[:15])
            print(f"[提示] '{q}' 包含命中多张表: {names}")
            return None
        print(f"[错误] 在 Index.xlsx 中未找到表: '{q}'")
        return None
    # 去重（保持顺序）
    seen, uniq = set(), []
    for r in selected:
        k = str(r[1]).strip().lower()
        if k not in seen:
            seen.add(k)
            uniq.append(r)
    return uniq


def build_temp_index(header, type_rows, selected):
    wb = __import__("openpyxl").Workbook()
    ws = wb.active
    ws.title = "index"
    for row in [header] + type_rows + selected:
        ws.append(list(row))
    wb.save(temp_index())


def run_export(mode, names=None):
    """执行 tabtoy 导出。"""
    out = out_dir()
    os.makedirs(os.path.join(out, "client"), exist_ok=True)
    os.makedirs(os.path.join(out, "client_test"), exist_ok=True)
    os.makedirs(os.path.join(out, "server"), exist_ok=True)
    os.makedirs(os.path.join(out, "server_test"), exist_ok=True)

    # 导出前清理旧的分表 json，避免被占用导致覆盖失败
    for name in (names or []):
        for d in ("client_test", "server_test"):
            old = os.path.join(out, d, f"{name}.json")
            if os.path.exists(old):
                try:
                    os.remove(old)
                    print(f"[清理] 删除旧文件 {d}/{name}.json")
                except OSError as e:
                    print(f"[提示] 无法删除旧文件 {d}/{name}.json: {e}")

    if mode in ("c", "cs"):
        print("[客户端] 导出中...")
        cmd = [TABTOY, "-mode=v3", f"-index={temp_index()}",
               "-json_dir=./out/client_test",
               "-binary_dir=./out/client",
               "-csharp_out=./out/client/config_partial.cs",
               "-tag_action=nogentab:server",
               "-package=Game.Hotfix.Config",
               "-tag_action=nogentab:server|nogenfield_lua:server"]
        r = subprocess.run(cmd, cwd=WORK_DIR)
        if r.returncode != 0:
            return r.returncode

    if mode in ("s", "cs"):
        print("[服务端] 导出中...")
        cmd = [TABTOY, "-mode=v3", f"-index={temp_index()}",
               "-json_out=./out/server/gameconfig_partial.json",
               "-json_dir=./out/server_test",
               "-proto_out=./out/server/gameconfig_partial.proto",
               "-tag_action=nogentab:client|nogenfield_json:client|nogenfield_pbbin:client"]
        r = subprocess.run(cmd, cwd=WORK_DIR)
        if r.returncode != 0:
            return r.returncode
    return 0


def list_tables(data_rows, kw=None):
    for r in data_rows:
        name = str(r[1]).strip()
        if kw and kw.lower() not in name.lower():
            continue
        desc = str(r[4]).strip() if len(r) > 4 and r[4] else ""
        tag = str(r[3]).strip() if len(r) > 3 and r[3] else ""
        line = f"{name:40s} tag={tag or '-':6s} {desc}"
        print(line)


def main(argv):
    if not argv or argv[0] in ("-h", "--help", "/?"):
        print(__doc__)
        return 0

    global WORK_DIR
    if argv[0] == "--dir":
        if len(argv) < 2:
            print("--dir 需要工作目录参数")
            return 1
        WORK_DIR = argv[1]
        argv = argv[2:]

    header, type_rows, data_rows = load_index()

    if argv and argv[0] == "--list":
        kw = argv[1] if len(argv) > 1 else None
        list_tables(data_rows, kw)
        return 0

    # 解析表名与选项
    queries, mode = [], "cs"
    for a in argv:
        al = a.lower()
        if al in ("-c", "--client"):
            mode = "c"
        elif al in ("-s", "--server"):
            mode = "s"
        elif al in ("-cs", "--both"):
            mode = "cs"
        elif al == "-prefix":
            continue
        else:
            for part in a.replace(",", " ").replace(";", " ").split():
                queries.append(part)

    if not queries:
        print("未指定任何表名。用 --list 查看可选表。")
        return 1

    selected = resolve_tables(data_rows, queries)
    if selected is None:
        return 1

    names = [str(r[1]).strip() for r in selected]
    tags = [str(r[3]).strip() if len(r) > 3 and r[3] else "" for r in selected]
    print(f"工作目录: {WORK_DIR}")
    print(f"将导出 {len(selected)} 张表 ({'/'.join({'c':'客户端','s':'服务端','cs':'两端'}[mode])}):")
    for n, t in zip(names, tags):
        print(f"  - {n}  (标记: {t or '两端'})")

    build_temp_index(header, type_rows, selected)
    try:
        rc = run_export(mode, names)
    finally:
        if os.path.exists(temp_index()):
            os.remove(temp_index())
    if rc != 0:
        print("导出失败")
        return rc
    print(f"\n完成! 输出目录: {out_dir()}")
    print("提示: config_partial.cs / gameconfig_partial.json 为部分产物, 仅供验证。")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
