# -*- coding: utf-8 -*-
"""启动器：检查依赖并启动 app。

供 run.bat 调用。把中文提示与依赖检查放到 Python 里，
避免 .bat 文件因中文编码在 cmd 下乱码/解析错误。
"""
import sys
import subprocess


def _has(import_name):
    try:
        __import__(import_name)
        return True
    except ImportError:
        return False


def main():
    print("=" * 52)
    print(" 配置表部分导出工具 - 启动中")
    print("=" * 52)

    # 必需依赖: openpyxl(读 Index.xlsx)
    if not _has("openpyxl"):
        print("检测到缺少依赖 openpyxl，正在自动安装...")
        r = subprocess.call([sys.executable, "-m", "pip", "install", "openpyxl"])
        if r != 0:
            print("[错误] openpyxl 安装失败，请手动执行:")
            print("        python -m pip install openpyxl")
            return 1
        print("openpyxl 安装完成。")

    # 可选依赖: pywebview(缺失时 app 会自动回退浏览器模式)
    if not _has("webview"):
        print("提示: 未安装 pywebview，将使用浏览器模式。")
        print("      如需独立窗口，可执行: python -m pip install pywebview")

    try:
        import app
        app.main()
    except Exception as e:
        print(f"[错误] 启动失败: {e}")
        return 1
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        pass
