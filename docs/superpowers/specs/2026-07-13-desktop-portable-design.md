# Excel 拆分工具桌面便携版设计

## 目标

将现有 Flask Excel 拆分工具封装为 Windows 10/11 免安装便携版。用户无需安装 Python，解压后双击 `Excel拆分工具.exe`，在无地址栏的独立窗口中使用现有功能。

## 范围

- 保留现有 Flask、HTML、CSS、JavaScript 和拆分引擎。
- 使用 pywebview 提供 Windows 独立窗口。
- 使用 Waitress 在 `127.0.0.1` 的动态端口运行本地服务。
- 窗口关闭后停止本地服务并退出进程。
- 使用 Windows 命名互斥量限制为单实例；再次启动时显示提示并退出。
- 将默认输出、日志和临时数据放在用户可写目录，不写入 PyInstaller 内部资源目录。
- 使用 PyInstaller `onedir`、`windowed` 生成便携目录。
- 提供 PowerShell 构建脚本和内部发布说明。

本轮不包含 Excel 拆分算法优化、全局任务队列、自动更新、安装器、公开分发和代码签名。

## 架构

`desktop.py` 是桌面入口。它获取单实例锁，创建运行目录和日志，启动 Waitress，等待首页可访问，再创建 pywebview 窗口。窗口事件循环结束后关闭 Waitress、释放锁并清理临时上传目录。

Flask 应用继续负责业务 API。应用工厂接收桌面入口提供的 `DEFAULT_OUTPUT_DIR` 和 `UPLOAD_DIR`，模板与静态资源仍作为 Python 包数据由 PyInstaller 收集。

## 数据目录

- 默认输出：`%USERPROFILE%\Documents\Excel拆分工具输出`
- 日志：`%LOCALAPPDATA%\ExcelSplitter\logs\excel-splitter.log`
- 上传临时目录：系统临时目录中的 `excel-splitter-*`，退出时清理

默认输出目录不得位于程序安装目录或 PyInstaller `_internal` 目录。

## 启动与关闭

1. 获取命名互斥量 `Local\ExcelSplitterTool`。
2. 初始化滚动日志。
3. 创建 Flask 应用和 Waitress 服务，绑定 `127.0.0.1:0`。
4. 后台线程运行 Waitress，轮询首页直到可访问。
5. 创建 1180x800、最小 900x640 的 pywebview 窗口。
6. 窗口关闭后调用服务关闭方法、等待线程结束并释放互斥量。

启动失败时记录完整异常，并通过 Windows 消息框显示中文错误，不依赖控制台。

## 打包

- 运行环境：CPython 3.11 x64。
- 运行依赖：Flask 3.1.3、openpyxl 3.1.5、pywebview 6.2.1、Waitress 3.0.2。
- 构建依赖：PyInstaller 6.21.0。
- 产物模式：`onedir` + `windowed`，支持文件放入 `_internal`。
- 发布目录包含主程序、内部运行库、`使用说明.txt` 和版本信息。

## 验证

- 单元测试桌面路径、端口、服务启动关闭和单实例锁的可测试边界。
- 运行现有全量测试，确保 Web 与拆分功能无回归。
- 构建便携目录并从 `dist` 启动 exe。
- 验证首页、文件上传、目录选择、关闭窗口后进程退出。
- 最终还需在一台未安装 Python 的 Windows 10/11 电脑执行干净机验收。

