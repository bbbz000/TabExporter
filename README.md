# TabExporter

> A GUI tool for selectively exporting & deploying game config tables, powered by [tabtoy](https://github.com/davyxu/tabtoy).

一个基于 [tabtoy](https://github.com/davyxu/tabtoy) 的游戏配置表**选择性导出 + 部署** GUI 工具。

不用每次全量导出几百张表——在图形界面里挑出要改的那 1~2 张表，一键导出，并可一键部署到 Unity 客户端工程。

## 特性

- **双列表交互**：左侧全部配置表（可搜索），右侧待导出表，点击/双击添加移除
- **选择性导出**：只导出选中的表，生成客户端 `.bin` + C#、服务端 JSON + Proto
- **工作目录可切换**：工具与「配置表目录」分离，可对任意配置表目录操作
- **一键部署**：把 `.bin` 复制为 `.bytes` 到 Unity 客户端（`config.cs` 可选）
- **双运行模式**：装了 `pywebview` 弹独立窗口，否则自动用浏览器打开
- **路径记忆**：工作目录、项目路径自动保存，下次启动带出

## 目录结构

```
TabExporter/
├── app.py              后端主程序（工作目录 / 导出 / 移动 / 双运行壳）
├── export_tables.py    核心库（读 Index.xlsx、构造临时 Index、调 tabtoy）
├── launcher.py         启动器（检查依赖后启动 app）
├── gui/
│   └── index.html      前端界面（单文件，原生 JS）
├── tabtoy.exe          第三方配置表导出工具（davyxu/tabtoy）
├── run.bat             双击运行（源码模式）
├── build_exe.bat       打包成单文件 exe
├── requirements.txt    依赖清单
└── tool_config.json    本地配置（运行时生成，已被 gitignore）
```

## 快速开始

### 运行（源码模式）

1. 安装 Python 3.10+（勾选 "Add to PATH"）
2. 双击 `run.bat`
3. 首次运行会自动安装 `openpyxl`（需联网）
4. 界面顶部点「切换工作目录…」，选择配置表目录（含 `Index.xlsx` 的目录）

> 未安装 `pywebview` 时自动用浏览器打开，功能完全一致；如需独立窗口：
> `python -m pip install pywebview`

也可命令行直接指定工作目录：

```bash
python launcher.py "C:\path\to\config\res01"
```

### 打包成 exe

双击 `build_exe.bat`，产物在 `dist\ExportTables.exe`。

> 目标电脑需已装 WebView2 Runtime（Win10/11 一般自带）。

## 使用说明

| 区域 | 说明 |
|---|---|
| 工作目录 | 顶部显示当前配置表目录，点「切换工作目录…」更换 |
| 左侧列表 | 全部配置表，顶部可搜索；单击选中，Ctrl+点击多选，双击直接添加 |
| 中间按钮 | ➜ 添加选中 · ⏩ 添加全部(当前筛选) · ⏪ 移除选中 · ✖ 清空 |
| 右侧列表 | 待导出表，双击移除 |
| 模式 | 两端 / 仅客户端 / 仅服务端，点「开始导出」 |
| 部署栏 | 填「项目路径」（客户端项目根目录），勾选「移动 config.cs」后点「移动」 |
| 日志 | 底部实时显示导出/移动日志 |

移动时的相对目录（相对「项目路径」）：

- 数据表：`Client\Project_LDL\Assets\ResPackage\DataTables`（`.bin` → `.bytes`）
- 代码：`Client\Project_LDL\Assets\Scripts\HotUpdate\GameFramework\LDLTable`（`config.cs`）

## 命令行版

核心库 `export_tables.py` 也可独立使用：

```bash
python export_tables.py --dir <工作目录> hero_config item_config   # 两端
python export_tables.py --dir <工作目录> hero_config -c            # 仅客户端
python export_tables.py --dir <工作目录> --list                    # 列出全部表
```

## 产物说明

- 客户端：覆盖对应表的 `out/client/*.bin`（增量），并生成 `out/client/config_partial.cs`
- 服务端：生成 `out/server/gameconfig_partial.json` / `.proto`，**不会**覆盖全量 `gameconfig.json`
- `config_partial.cs` / `gameconfig_partial.json` 仅供验证查错，不要替代全量 `config.cs` / `gameconfig.json` 部署
- 带 `client` / `server` 标记的表，导出另一端时会被 tabtoy 自动跳过（属正常）

## 第三方依赖

- [tabtoy](https://github.com/davyxu/tabtoy)（MIT）—— 配置表导出核心，本仓库内置其 `tabtoy.exe` 用于运行
- [openpyxl](https://openpyxl.readthedocs.io/)（MIT）—— 读取 `Index.xlsx`
- [pywebview](https://pywebview.flowrl.com/)（BSD-3-Clause，可选）—— 桌面窗口壳

## License

本项目代码按 [MIT License](LICENSE) 发布（若你尚未添加 LICENSE 文件，可自行选择协议）。
