# -*- coding: utf-8 -*-
"""
配置表部分导出工具 —— 后端
运行方式:
  python app.py [工作目录]   # 优先 pywebview 原生窗口，无 pywebview 则自动开浏览器
前端: gui/index.html

工作目录 = 存放 Index.xlsx / *.xlsx / out/ 的配置表目录，可在界面里切换。
依赖: openpyxl (读 Index.xlsx)，可选 pywebview (桌面窗口壳)。
"""
import os
import sys
import json
import subprocess
import threading
import webbrowser
import shutil

TOOL_DIR = os.path.dirname(os.path.abspath(__file__))


def _data_dir():
    """持久数据目录：打包成 exe 后用 exe 所在目录(避免 _MEIPASS 临时目录丢失)，源码模式用脚本目录。"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return TOOL_DIR


CONFIG_FILE = os.path.join(_data_dir(), "tool_config.json")

# 复用 export_tables.py 的纯逻辑函数
sys.path.insert(0, TOOL_DIR)
import export_tables as et  # noqa: E402

WORK_DIR = ""  # 运行时确定的工作目录


def resource_path(rel):
    """打包成 exe 后从 _MEIPASS 定位资源，否则用工具目录。"""
    base = getattr(sys, "_MEIPASS", TOOL_DIR)
    return os.path.join(base, rel)


def is_valid_work_dir(path):
    path = (path or "").strip().strip('"')
    return bool(path) and os.path.isdir(path) and os.path.isfile(os.path.join(path, "Index.xlsx"))


def get_config():
    """读取持久化配置(页签工作目录 / 项目路径)。"""
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    except Exception:
        cfg = {}
    # 向后兼容：旧版 work_dir 字段迁移为 work_dirs + active_work_dir
    if "work_dirs" not in cfg:
        old = cfg.get("work_dir", "")
        cfg["work_dirs"] = [old] if old else []
        cfg["active_work_dir"] = old
        cfg.pop("work_dir", None)
    if "active_work_dir" not in cfg:
        dirs = cfg.get("work_dirs", [])
        cfg["active_work_dir"] = dirs[0] if dirs else ""
    # 规范化路径，避免同一目录因正/反斜杠不同而被重复添加
    cfg["work_dirs"] = list(dict.fromkeys(
        os.path.normpath(d) for d in cfg.get("work_dirs", []) if d
    ))
    if cfg.get("active_work_dir"):
        cfg["active_work_dir"] = os.path.normpath(cfg["active_work_dir"])
    return cfg


def save_config(work_dirs=None, active_work_dir=None, project_path=None):
    cfg = get_config()
    if work_dirs is not None:
        cfg["work_dirs"] = work_dirs
    if active_work_dir is not None:
        cfg["active_work_dir"] = active_work_dir
    if project_path is not None:
        cfg["project_path"] = project_path
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


def get_work_dir():
    return WORK_DIR


def get_tabs():
    """返回所有页签(工作目录)和当前激活目录。"""
    cfg = get_config()
    dirs = [d for d in cfg.get("work_dirs", []) if d]
    return {"tabs": dirs, "active": WORK_DIR or cfg.get("active_work_dir", "")}


def set_work_dir(path):
    """切换工作目录(自动加入页签)并持久化。返回 {ok, log, tabs, active}。"""
    global WORK_DIR
    path = os.path.normpath((path or "").strip().strip('"'))
    if not is_valid_work_dir(path):
        return {"ok": False, "log": f"无效工作目录(需包含 Index.xlsx): {path}"}
    WORK_DIR = path
    et.WORK_DIR = path
    cfg = get_config()
    dirs = [d for d in cfg.get("work_dirs", []) if d]
    if path not in dirs:
        dirs.append(path)
    save_config(work_dirs=dirs, active_work_dir=path)
    return {"ok": True, "log": f"工作目录已切换: {path}", "tabs": dirs, "active": path}


def remove_tab(path):
    """移除一个页签；若移除的是当前激活页签，则切到剩余第一个。"""
    global WORK_DIR
    path = os.path.normpath((path or "").strip().strip('"'))
    cfg = get_config()
    dirs = [d for d in cfg.get("work_dirs", []) if d and d != path]
    active = cfg.get("active_work_dir", "")
    if active == path or active not in dirs:
        active = dirs[0] if dirs else ""
        WORK_DIR = active
        et.WORK_DIR = active
    save_config(work_dirs=dirs, active_work_dir=active)
    return {"ok": True, "log": f"已移除页签: {path}", "tabs": dirs, "active": active}


def list_tables():
    """读取工作目录下的 Index.xlsx，返回所有数据表。"""
    if not WORK_DIR:
        return []
    et.WORK_DIR = WORK_DIR
    _, _, data_rows = et.load_index()
    tables = []
    for r in data_rows:
        name = str(r[1]).strip()
        file_ = str(r[2]).strip() if len(r) > 2 and r[2] else ""
        tag = str(r[3]).strip() if len(r) > 3 and r[3] else ""
        desc = str(r[4]).strip() if len(r) > 4 and r[4] else ""
        tables.append({"name": name, "file": file_, "tag": tag, "desc": desc})
    return tables


def _decode(b):
    """tabtoy 输出编码可能是 utf-8 或 gbk，逐种尝试解码。"""
    if not b:
        return ""
    for enc in ("utf-8", "gbk"):
        try:
            return b.decode(enc)
        except UnicodeDecodeError:
            continue
    return b.decode("utf-8", errors="replace")


def _run_tabtoy(cmd, cwd):
    """执行 tabtoy，尽量捕获日志；环境禁用管道捕获时降级为直连控制台。"""
    try:
        r = subprocess.run(cmd, cwd=cwd, capture_output=True)
        log = _decode(r.stdout) + _decode(r.stderr)
        return r.returncode == 0, log
    except OSError as e:
        r = subprocess.run(cmd, cwd=cwd)
        return r.returncode == 0, f"(当前环境不支持捕获日志，已直接输出到控制台)\n{e}\n"


def export_tables(names, mode):
    """部分导出：构造精简 Index 后调用 tabtoy。返回 {ok, log}。"""
    if not WORK_DIR:
        return {"ok": False, "log": "未设置工作目录"}
    if not names:
        return {"ok": False, "log": "未指定任何表"}
    mode = mode if mode in ("c", "s", "cs") else "cs"
    et.WORK_DIR = WORK_DIR

    try:
        header, type_rows, data_rows = et.load_index()
    except Exception as e:
        return {"ok": False, "log": f"读取 Index.xlsx 失败: {e}"}

    selected = et.resolve_tables(data_rows, list(names))
    if selected is None:
        return {"ok": False, "log": "存在未找到的表名，已中止"}

    sel_names = [str(r[1]).strip() for r in selected]
    log_lines = [f"工作目录: {WORK_DIR}",
                 f"将导出 {len(sel_names)} 张表: {', '.join(sel_names)}",
                 f"模式: {mode}"]

    try:
        et.build_temp_index(header, type_rows, selected)
    except Exception as e:
        return {"ok": False, "log": f"生成临时 Index 失败: {e}"}

    ok = True
    try:
        for d in ("client", "client_test", "server", "server_test"):
            os.makedirs(os.path.join(WORK_DIR, "out", d), exist_ok=True)

        # 导出前清理旧的分表 json，避免被查看器占用导致覆盖失败
        for name in sel_names:
            for d in ("client_test", "server_test"):
                old = os.path.join(WORK_DIR, "out", d, f"{name}.json")
                if os.path.exists(old):
                    try:
                        os.remove(old)
                        log_lines.append(f"[清理] 删除旧文件 {d}/{name}.json")
                    except OSError as e:
                        log_lines.append(f"[提示] 无法删除旧文件 {d}/{name}.json: {e}")

        if mode in ("c", "cs"):
            log_lines.append("[客户端] 导出中...")
            cmd = [et.TABTOY, "-mode=v3", f"-index={et.temp_index()}",
                   "-json_dir=./out/client_test",
                   "-binary_dir=./out/client",
                   "-csharp_out=./out/client/config_partial.cs",
                   "-tag_action=nogentab:server",
                   "-package=Game.Hotfix.Config",
                   "-tag_action=nogentab:server|nogenfield_lua:server"]
            okc, lg = _run_tabtoy(cmd, cwd=WORK_DIR)
            log_lines.append(lg)
            ok = ok and okc

        if mode in ("s", "cs"):
            log_lines.append("[服务端] 导出中...")
            cmd = [et.TABTOY, "-mode=v3", f"-index={et.temp_index()}",
                   "-json_out=./out/server/gameconfig_partial.json",
                   "-json_dir=./out/server_test",
                   "-proto_out=./out/server/gameconfig_partial.proto",
                   "-tag_action=nogentab:client|nogenfield_json:client|nogenfield_pbbin:client"]
            oks, lg = _run_tabtoy(cmd, cwd=WORK_DIR)
            log_lines.append(lg)
            ok = ok and oks
    finally:
        tmp = et.temp_index()
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass

    log_lines.append("✅ 完成" if ok else "❌ 失败")
    return {"ok": ok, "log": "\n".join(log_lines)}


# ================= 移动/部署 =================
REL_DATATABLES = os.path.join("Client", "Project_LDL", "Assets", "ResPackage", "DataTables")
REL_CODE = os.path.join("Client", "Project_LDL", "Assets", "Scripts", "HotUpdate", "GameFramework", "LDLTable")


def choose_dir():
    """弹出目录选择框，返回所选绝对路径；失败返回空串。"""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.askdirectory(title="选择目录")
        try:
            root.destroy()
        except Exception:
            pass
        return path or ""
    except Exception:
        return ""


def move_tables(names, project_path, move_cs=False):
    """把本次导出的表复制到客户端工程。返回 {ok, log}。"""
    if not WORK_DIR:
        return {"ok": False, "log": "未设置工作目录"}
    if not names:
        return {"ok": False, "log": "未指定任何表"}
    project_path = (project_path or "").strip().strip('"')
    if not project_path:
        return {"ok": False, "log": "请先设置项目路径"}
    if not os.path.isdir(project_path):
        return {"ok": False, "log": f"项目路径不存在: {project_path}"}

    dt_dir = os.path.join(project_path, REL_DATATABLES)
    code_dir = os.path.join(project_path, REL_CODE)
    logs = [f"工作目录: {WORK_DIR}",
            f"项目路径: {project_path}",
            f"数据表目录: {dt_dir}",
            f"代码目录: {code_dir}"]
    try:
        os.makedirs(dt_dir, exist_ok=True)
        os.makedirs(code_dir, exist_ok=True)
    except OSError as e:
        return {"ok": False, "log": f"创建目标目录失败: {e}"}

    moved = failed = skipped = 0
    for name in names:
        src = os.path.join(WORK_DIR, "out", "client", f"{name}.bin")
        dst = os.path.join(dt_dir, f"{name}.bytes")
        if not os.path.exists(src):
            logs.append(f"[跳过] 未找到 {name}.bin（请先导出该表）")
            skipped += 1
            continue
        try:
            shutil.copyfile(src, dst)
            moved += 1
            logs.append(f"[复制] {name}.bin -> {name}.bytes")
        except OSError as e:
            failed += 1
            logs.append(f"[失败] {name}: {e}")

    if move_cs:
        cs_src = os.path.join(WORK_DIR, "out", "client", "config.cs")
        if os.path.exists(cs_src):
            try:
                shutil.copyfile(cs_src, os.path.join(code_dir, "config.cs"))
                logs.append("[复制] config.cs -> 代码目录")
            except OSError as e:
                logs.append(f"[失败] config.cs: {e}")
        else:
            logs.append("[提示] 未找到 out/client/config.cs（全量导出才有）")
    else:
        logs.append("[跳过] config.cs（未勾选移动 config.cs）")

    save_config(project_path=project_path)
    ok = failed == 0
    logs.append(f"完成: 成功 {moved} 个，跳过 {skipped} 个，失败 {failed} 个")
    return {"ok": ok, "log": "\n".join(logs)}


# ================= 运行壳 =================
HTML_PATH = resource_path(os.path.join("gui", "index.html"))


def run_webview():
    import webview

    class Api:
        def list_tables(self):
            return list_tables()

        def export_tables(self, names, mode):
            return export_tables(names, mode)

        def move_tables(self, names, project_path, move_cs):
            return move_tables(names, project_path, move_cs)

        def get_config(self):
            return get_config()

        def get_work_dir(self):
            return get_work_dir()

        def set_work_dir(self, path):
            return set_work_dir(path)

        def get_tabs(self):
            return get_tabs()

        def remove_tab(self, path):
            return remove_tab(path)

        def choose_dir(self):
            return choose_dir()

    webview.create_window("配置表部分导出工具", HTML_PATH,
                          js_api=Api(), width=1180, height=780,
                          min_size=(900, 580))
    webview.start()


def _api_handler():
    from http.server import BaseHTTPRequestHandler

    class Handler(BaseHTTPRequestHandler):
        def _send(self, code, body, ctype="application/json; charset=utf-8"):
            data = body if isinstance(body, bytes) else body.encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)

        def do_GET(self):
            if self.path in ("/", "/index.html"):
                with open(HTML_PATH, "rb") as f:
                    self._send(200, f.read(), "text/html; charset=utf-8")
            elif self.path == "/api/list":
                self._send(200, json.dumps(list_tables(), ensure_ascii=False))
            elif self.path == "/api/config":
                self._send(200, json.dumps(get_config(), ensure_ascii=False))
            elif self.path == "/api/work_dir":
                self._send(200, json.dumps({"work_dir": get_work_dir()}, ensure_ascii=False))
            elif self.path == "/api/tabs":
                self._send(200, json.dumps(get_tabs(), ensure_ascii=False))
            else:
                self._send(404, "not found")

        def do_POST(self):
            if self.path == "/api/export":
                n = int(self.headers.get("Content-Length", 0))
                payload = json.loads(self.rfile.read(n) or b"{}")
                res = export_tables(payload.get("names", []), payload.get("mode", "cs"))
                self._send(200, json.dumps(res, ensure_ascii=False))
            elif self.path == "/api/move":
                n = int(self.headers.get("Content-Length", 0))
                payload = json.loads(self.rfile.read(n) or b"{}")
                res = move_tables(payload.get("names", []), payload.get("project_path", ""), payload.get("move_cs", False))
                self._send(200, json.dumps(res, ensure_ascii=False))
            elif self.path == "/api/set_work_dir":
                n = int(self.headers.get("Content-Length", 0))
                payload = json.loads(self.rfile.read(n) or b"{}")
                res = set_work_dir(payload.get("path", ""))
                self._send(200, json.dumps(res, ensure_ascii=False))
            elif self.path == "/api/choose_dir":
                self._send(200, json.dumps({"path": choose_dir()}, ensure_ascii=False))
            elif self.path == "/api/remove_tab":
                n = int(self.headers.get("Content-Length", 0))
                payload = json.loads(self.rfile.read(n) or b"{}")
                res = remove_tab(payload.get("path", ""))
                self._send(200, json.dumps(res, ensure_ascii=False))
            else:
                self._send(404, "not found")

        def log_message(self, *a):
            pass

    return Handler


def run_browser(port=0):
    from http.server import HTTPServer

    Handler = _api_handler()
    server = HTTPServer(("127.0.0.1", port), Handler)
    port = server.server_address[1]
    url = f"http://127.0.0.1:{port}/"
    print(f"服务已启动: {url}", flush=True)
    threading.Timer(0.4, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


def main():
    global WORK_DIR
    cfg = get_config()
    work_dir = cfg.get("active_work_dir", "")
    if len(sys.argv) > 1:
        work_dir = sys.argv[1]
    if is_valid_work_dir(work_dir):
        set_work_dir(work_dir)
    # 否则 WORK_DIR 保持空，界面显示"未设置工作目录"，由用户切换

    # run.bat 通过 TABEXPORTER_NO_WEBVIEW=1 强制浏览器(源码模式用系统 Python 3.14 + pywebview 6.x 有 bug)；
    # exe 默认用 pywebview 独立窗口(内置 Python 3.12 + pywebview 5.4)。
    if os.environ.get('TABEXPORTER_NO_WEBVIEW') == '1':
        run_browser()
        return
    try:
        import webview  # noqa: F401
        run_webview()
    except ImportError:
        run_browser()


if __name__ == "__main__":
    main()
