# TabExporter 二次开发文档

面向希望**二次开发、维护、扩展**本工具的开发者或 AI 助手。本文档记录工程架构、代码结构、数据流、关键机制与已知坑，帮助你快速上手改动。

---

## 1. 工程概览

TabExporter 是一个「配置表选择性导出 + 部署」桌面工具，技术栈非常简单：

| 层 | 技术 | 说明 |
|---|---|---|
| 前端 | 原生 HTML/CSS/JS（单文件） | `gui/index.html`，无框架、无构建步骤 |
| 后端 | Python 3（标准库 + openpyxl） | `app.py` + `export_tables.py` |
| 导出核心 | tabtoy.exe（第三方 Go 工具） | 负责把 Excel 转成 bin/json/csharp/proto |
| 桌面壳 | pywebview 5.x | exe 独立窗口；源码模式用浏览器 |

**一句话架构**：前端（单页）通过统一 API 层调用后端 Python 函数，后端构造「精简 Index.xlsx」后调用 `tabtoy.exe` 完成导出。

---

## 2. 目录结构

```
TabExporter/
├── app.py              后端主程序（多页签 / 导出 / 移动 / 双运行壳）
├── export_tables.py    核心库（纯逻辑，可命令行单独运行）
├── launcher.py         启动器（依赖检查 + 启动 app）
├── gui/index.html      前端界面（单文件）
├── tabtoy.exe          tabtoy 二进制（第三方）
├── run.bat / build_exe.bat
├── requirements.txt    pywebview 锁定 5.x
└── tool_config.json    本地配置（运行时生成，gitignore）
```

---

## 3. 核心概念

### 3.1 工具目录 vs 工作目录

- **工具目录（TOOL_DIR）**：`app.py` / `export_tables.py` 所在目录，存放工具代码 + `tabtoy.exe` + `gui/`。
- **工作目录（WORK_DIR）**：存放 `Index.xlsx`、`*.xlsx`、`out/` 的**配置表目录**（例如 `res01`）。由用户指定，可多页签切换。

### 3.2 多页签（多工作目录）

配置表里持久化**多个工作目录**，每个是一个页签：

```json
{
  "work_dirs": [".../dev01", ".../res01"],
  "active_work_dir": ".../dev01",
  "project_path": "..."
}
```

- `set_work_dir(path)`：切换工作目录，并**自动加入** `work_dirs`（路径经 `normpath` 规范化 + 去重，避免正/反斜杠重复）。
- `get_tabs()`：返回 `{tabs, active}`。
- `remove_tab(path)`：移除页签；若移除的是当前激活页签，切到剩余第一个。
- 前端每个页签独立保存选择状态到 `tabStates`（内存 + `localStorage` 持久化）。

### 3.3 双运行模式

`app.py` 的 `main()` 决定运行模式：

- 环境变量 `TABEXPORTER_NO_WEBVIEW=1`（`run.bat` 设置）→ `run_browser()`：`http.server` + `fetch('/api/xxx')`。
- 否则（含打包 exe）→ `run_webview()`：pywebview 独立窗口，`pywebview.api.xxx()`。

> 前端通过 `api` 对象统一封装两种模式；启动时用 `window.chrome.webview` 判断是否 WebView2 环境，是则等 `pywebviewready` 事件再初始化（见 §11.6）。

---

## 4. 后端 app.py

### 4.1 全局状态与路径

| 变量 | 含义 |
|---|---|
| `TOOL_DIR` | 工具目录 |
| `_data_dir()` | 持久数据目录：打包 exe 后 = `sys.executable` 所在目录；源码 = `TOOL_DIR` |
| `CONFIG_FILE` | `_data_dir()/tool_config.json` |
| `WORK_DIR` | 当前激活的工作目录（空串表示未设置） |
| `et` | import 进来的 `export_tables` 模块 |

> 关键：exe（PyInstaller onefile）下 `__file__` 指向临时 `_MEIPASS`，退出即删，所以 `CONFIG_FILE` 必须用 `_data_dir()`（exe 目录）持久化，否则配置每次丢失。

### 4.2 函数清单

| 函数 | 作用 | 返回 |
|---|---|---|
| `resource_path(rel)` | 打包后从 `_MEIPASS` 定位资源 | 绝对路径 |
| `is_valid_work_dir(path)` | 校验目录含 `Index.xlsx` | bool |
| `get_config()` | 读 `tool_config.json`，含旧 `work_dir` 迁移 + 路径规范化 | dict |
| `save_config(work_dirs, active_work_dir, project_path)` | 持久化配置 | None |
| `get_work_dir()` / `set_work_dir(path)` | 读/切工作目录（自动加页签） | str / `{ok, log, tabs, active}` |
| `get_tabs()` / `remove_tab(path)` | 列页签 / 移除页签 | `{tabs, active}` / `{ok, log, tabs, active}` |
| `list_tables()` | 读工作目录 `Index.xlsx`，返回全部数据表 | `[{name,file,tag,desc}]` |
| `_decode(b)` | 解码 tabtoy 输出（utf-8 → gbk 逐种尝试） | str |
| `_run_tabtoy(cmd, cwd)` | 执行 tabtoy 并捕获日志 | `(ok, log)` |
| `export_tables(names, mode)` | 部分导出（核心业务） | `{ok, log}` |
| `choose_dir()` | tkinter 弹目录选择框 | 路径或空串 |
| `move_tables(names, project_path, move_cs)` | 复制 bin→bytes 到客户端 | `{ok, log}` |
| `run_webview()` / `run_browser()` / `main()` | 运行壳 | — |

### 4.3 常量

```python
REL_DATATABLES = "Client\\Project_LDL\\Assets\\ResPackage\\DataTables"
REL_CODE       = "Client\\Project_LDL\\Assets\\Scripts\\HotUpdate\\GameFramework\\LDLTable"
```

这两个是「项目路径」下的相对子目录，改动移动行为时改这里即可。

---

## 5. 核心库 export_tables.py

### 5.1 全局

```python
SCRIPT_DIR  # 工具目录（tabtoy.exe 所在）
TABTOY      # SCRIPT_DIR/tabtoy.exe
WORK_DIR    # 工作目录，默认 os.getcwd()，app.py 会覆盖
```

路径通过函数动态获取（而非模块级常量），以保证 `WORK_DIR` 切换后生效：

```python
index_file()   # WORK_DIR/Index.xlsx
temp_index()   # WORK_DIR/_temp_index.xlsx
out_dir()      # WORK_DIR/out
```

### 5.2 函数

| 函数 | 作用 |
|---|---|
| `load_index()` | 打开 `Index.xlsx` 的 `index` sheet，返回 `(header, type_rows, data_rows)` |
| `resolve_tables(data_rows, queries)` | 表名匹配（精确 → 前缀 → 包含），返回选中的行；失败返回 `None` |
| `build_temp_index(header, type_rows, selected)` | 写精简临时 `_temp_index.xlsx`（表头 + 类型表行 + 选中行） |
| `run_export(mode, names=None)` | 清旧 json → 调 tabtoy（客户端/服务端两条命令） |
| `main(argv)` | 命令行入口，支持 `--dir <工作目录>`、`--list`、`-c/-s/-cs` |

### 5.3 关键点：精简 Index 机制

tabtoy 没有「只导某几张表」的参数，它只认 `-index` 指向的 Index 文件。所以部分导出 = **构造一个只含「2 张类型表 + 要导的 N 张数据表」的临时 Index**，再原样调 tabtoy。

- `type_rows`（类型表行）**必须保留**：`game_attribute.xlsx`、`game_enumeration.xlsx` 定义全局字段类型和枚举，数据表都引用它们。
- 临时 Index 用完即删（`temp_index()` 在 `finally` 里移除）。

---

## 6. 前端 gui/index.html

### 6.1 统一 API 层

前端 `api` 对象是唯一访问后端的入口，内部自动分流 pywebview / http：

```js
const api = {
  async getConfig(){
    if(window.pywebview && window.pywebview.api)
      return await window.pywebview.api.get_config();
    const r = await fetch('/api/config'); return await r.json();
  },
  // listTables / exportTables / moveTables / getTabs / removeTab / setWorkDir / chooseDir 同理
};
```

**加新功能时**：后端加函数 → pywebview Api 类加方法 → http handler 加路由 → 前端 `api` 对象加方法。

### 6.2 状态与渲染

```js
let allTables = [];      // 左列表全部
let rightSet  = [];      // 右侧待导出表名（有序）
let selLeft / selRight = new Set();  // 两侧选中项
let mode = 'c';          // 导出模式（默认仅客户端）
let workDir = '';        // 当前工作目录
let tabs = [];           // 页签列表（工作目录路径）
let tabStates = {};      // 每页签独立状态 {path: {rightSet, selLeft, selRight, search}}
```

- `renderTabs()` / `renderWorkDir()` / `renderLeft()` / `renderRight()` 渲染各区域
- `saveCurrentState()` / `restoreState(path)`：切页签时保存/恢复状态（`restoreState` 用**原地更新** `clear()+add()`，不能重新赋值 Set，否则 `bindList` 闭包持有的引用脱节）
- `log(msg, cls)` 写底部日志区

### 6.3 交互入口

- `bindList(...)`：单击选中、Ctrl 多选、**Shift 范围多选**、双击添加/移除
- `switchTab(path)` / `addTab()` / `removeTab(path)`：页签切换/新增/关闭
- `doExport()` / `doMove()`：导出 / 移动
- 空格键：对当前选中项快速移动（左→右添加 / 右→左移除）
- `initDivider()` / `initDividerH()`：左右分栏 / 日志栏高度拖动

### 6.4 启动时序

页面脚本末尾不直接调 `init()`，而是：

```js
if(window.chrome && window.chrome.webview){
  // WebView2(pywebview)：等 pywebviewready(js_api 注册完成) 再 init
  if(window.pywebview?.api && Object.keys(window.pywebview.api).length) _boot();
  else { window.addEventListener('pywebviewready', _boot, {once:true}); setTimeout(_boot, 3000); }
} else {
  _boot();  // 普通浏览器：直接 init
}
```

---

## 7. API 接口对照（pywebview ↔ http）

| 前端调用 | pywebview 方法 | HTTP 端点 |
|---|---|---|
| `api.listTables()` | `Api.list_tables` | `GET /api/list` |
| `api.exportTables(names, mode)` | `Api.export_tables` | `POST /api/export` |
| `api.moveTables(names, path, moveCs)` | `Api.move_tables` | `POST /api/move` |
| `api.getConfig()` | `Api.get_config` | `GET /api/config` |
| `api.getWorkDir()` | `Api.get_work_dir` | `GET /api/work_dir` |
| `api.setWorkDir(path)` | `Api.set_work_dir` | `POST /api/set_work_dir` |
| `api.getTabs()` | `Api.get_tabs` | `GET /api/tabs` |
| `api.removeTab(path)` | `Api.remove_tab` | `POST /api/remove_tab` |
| `api.chooseDir()` | `Api.choose_dir` | `POST /api/choose_dir` |

注意：pywebview 的方法名保持 Python 下划线风格（`get_tabs`），http 端点是斜杠风格。

---

## 8. 数据流

### 8.1 导出流程

```
用户点「开始导出」
  → doExport() 收集 rightSet + mode
  → api.exportTables(names, mode)
  → app.export_tables(names, mode)
      1. 校验 WORK_DIR / names / mode
      2. et.load_index()            读 Index.xlsx
      3. et.resolve_tables()        匹配选中表
      4. et.build_temp_index()      写精简 _temp_index.xlsx
      5. 删除 client_test/server_test 下旧的分表 json（防占用）
      6. _run_tabtoy(cmd, cwd=WORK_DIR)  调 tabtoy（客户端 + 服务端）
      7. 清理 _temp_index.xlsx
  → 返回 {ok, log} → 前端逐行打印日志
```

### 8.2 移动流程

```
用户点「移动」
  → doMove() 收集 rightSet + 项目路径 + chkMoveCs
  → api.moveTables(names, project_path, move_cs)
  → app.move_tables()
      1. 拼 dt_dir = 项目路径 + REL_DATATABLES
         code_dir = 项目路径 + REL_CODE
      2. 对每个表：WORK_DIR/out/client/{name}.bin
         → copy → dt_dir/{name}.bytes
      3. 若 move_cs：copy out/client/config.cs → code_dir/config.cs
  → 返回 {ok, log}
```

---

## 9. Index.xlsx 结构

`Index.xlsx` 的 sheet 名为 `index`，这是 tabtoy 的约定：

| 列 | 含义 |
|---|---|
| A | 模式（`类型表` / `数据表`） |
| B | 表类型（数据表的表名；类型表此列为空） |
| C | 表文件名（`.xlsx`） |
| D | 标记（`client` / `server` / 空 = 两端） |
| E | 描述/备注 |
| G | langid 起始值（仅部分表） |

判定规则（见 `load_index()`）：

```python
type_rows = [r for r in rows[1:] if r[0] and not r[1]]   # A非空 且 B空
data_rows = [r for r in rows[1:] if r[1]]                 # B非空
```

---

## 10. tabtoy 参数说明

客户端导出（`export_tables.py` / `app.py` 中）：

```
tabtoy.exe -mode=v3 -index=<临时Index>
  -json_dir=./out/client_test      # 分表 json（调试用）
  -binary_dir=./out/client         # 正式 .bin
  -csharp_out=./out/client/config_partial.cs
  -tag_action=nogentab:server      # 跳过带 server 标记的
  -package=Game.Hotfix.Config
  -tag_action=nogentab:server|nogenfield_lua:server
```

服务端导出：

```
tabtoy.exe -mode=v3 -index=<临时Index>
  -json_out=./out/server/gameconfig_partial.json   # 聚合 json
  -json_dir=./out/server_test                      # 分表 json
  -proto_out=./out/server/gameconfig_partial.proto
  -tag_action=nogentab:client|nogenfield_json:client|nogenfield_pbbin:client
```

- `-tag_action=nogentab:xxx`：对带 `xxx` 标记的表/字段**不生成**（用于按 client/server 分流）。
- tabtoy 运行 `cwd=WORK_DIR`，所以 `./out/...` 相对路径解析到工作目录下的 `out/`。

---

## 11. 已知坑与注意事项

### 11.1 分表 json 被占用导致覆盖失败

现象：`open ./out/client_test/xxx.json: Access is denied`，但 bin/cs 写成功、且该 json「读 OK、写失败、删除 OK」。

原因：旧 json 被某进程以「只读映射/只读共享」方式占用。

解法（已内置）：导出前先删除 `client_test`/`server_test` 下对应表的旧 json。

### 11.2 tabtoy 输出编码

tabtoy 输出可能是 GBK 或 UTF-8。`_decode()` 逐种尝试解码，**不要**用 `subprocess.run(..., text=True)` 默认编码。

### 11.3 .bat 文件中文乱码

`.bat` 若含中文且存为 UTF-8（无 BOM），中文 Windows 的 cmd 用 GBK 解析会错位报错。

解法（已采用）：**`.bat` 全部纯 ASCII 英文**，中文提示移到 Python（`launcher.py`）。

### 11.4 部分导出的代码不完整

部分导出只生成 `config_partial.cs`（只含导出表的结构），**不能**替换全量 `config.cs`。移动时 `config.cs` 默认不移动，需手动勾选。

### 11.5 沙箱环境的子进程管道

某些受限环境禁止捕获子进程 stdio（`capture_output` 报 EPERM）。`_run_tabtoy` 已做降级：捕获失败时改用继承 stdio。

### 11.6 pywebview 6.x js_api 时序 bug（重要）

现象：独立窗口下随机报 `getConfig 异常: SyntaxError: Unexpected token <!DOCTYPE...`，浏览器模式正常。

原因：pywebview 的 js_api（`window.pywebview`）是**异步注入**的，前端 `init()` 执行太早时 `window.pywebview` 尚未注入，前端误走 `fetch('/api/config')`，而 pywebview 内部的 bottle 服务器没有 `/api/config`，返回 404 HTML。

解法：
1. **pywebview 锁定 5.x**（`requirements.txt` 的 `pywebview>=5.0,<6`），6.2.1 更易触发；
2. 前端用 `window.chrome.webview` 判断是否 WebView2 环境，是则**等 `pywebviewready` 事件**再 `init()`（见 §6.4）。

### 11.7 exe 下配置丢失（_MEIPASS）

PyInstaller `--onefile` 下 `__file__` 指向临时解压目录 `_MEIPASS`，退出即删，导致写在那里的 `tool_config.json` 每次丢失。

解法：`CONFIG_FILE` 用 `_data_dir()`（`sys.frozen` 时取 `sys.executable` 所在目录），配置持久化到 exe 旁边。

### 11.8 路径正/反斜杠导致重复页签

同一目录因正斜杠（`C:/...`）与反斜杠（`C:\...`）不同会被判为两个页签。

解法：`get_config()` 里对 `work_dirs` 和 `active_work_dir` 统一 `os.path.normpath` 并去重。

---

## 12. 扩展指南

### 12.1 加一个新的后端操作（以「导出到额外格式」为例）

1. `app.py` 加函数 `do_xxx(...)`，返回 `{ok, log}`；
2. `run_webview()` 的 `Api` 类加 `def xxx(...)`；
3. `_api_handler()` 的 `do_POST` 加 `elif self.path == "/api/xxx":`；
4. `gui/index.html` 的 `api` 对象加 `async xxx(...)`；
5. 前端加按钮 + 事件绑定。

### 12.2 改部署相对目录

改 `app.py` 的 `REL_DATATABLES` / `REL_CODE` 两个常量即可。

### 12.3 增加命令行参数

`export_tables.py` 的 `main(argv)` 是纯手写参数解析，按现有 `-c/-s/--dir/--list` 模式加即可。

---

## 13. 打包说明

`build_exe.bat` 使用 PyInstaller：

```
pyinstaller --onefile --windowed --name ExportTables \
  --add-data "gui;gui"       # 前端资源
  --add-data "tabtoy.exe;."  # tabtoy 二进制
  --collect-all webview      # pywebview 依赖
  app.py
```

打包后 `sys._MEIPASS` 指向解压临时目录，`resource_path()` 和 `SCRIPT_DIR` 会自动适配（`__file__` 指向 `_MEIPASS`），因此 `tabtoy.exe` 和 `gui/` 都能被正确找到；但**持久配置**（`tool_config.json`）必须用 `_data_dir()`（exe 目录），见 §11.7。
