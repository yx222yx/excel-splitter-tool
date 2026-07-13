# Excel 拆分工具桌面便携版 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 生成无需 Python、解压后双击即可在独立窗口运行的 Windows 便携版。

**Architecture:** 新增桌面启动层，在动态回环端口运行 Waitress，并由 pywebview 承载现有 Flask UI。PyInstaller `onedir` 收集 Python 运行时和 Web 静态资源，用户数据写入 Documents 与 LocalAppData。

**Tech Stack:** Python 3.11、Flask 3.1.3、openpyxl 3.1.5、Waitress 3.0.2、pywebview 6.2.1、PyInstaller 6.21.0。

## Global Constraints

- 仅支持 Windows 10/11 x64 内部分发。
- 用户无需安装 Python。
- 第一版为免安装便携目录，不做安装器和自动更新。
- 保留现有 Web UI 与拆分引擎，不重构业务逻辑。
- 输出和日志不得写入 PyInstaller 资源目录。

---

### Task 1: 桌面运行时边界

**Files:**
- Create: `src/excel_splitter/desktop_runtime.py`
- Test: `tests/test_desktop_runtime.py`

**Interfaces:**
- Produces: `user_output_dir() -> Path`、`log_file_path() -> Path`、`SingleInstance`、`wait_until_ready(url, timeout)`。

- [ ] 写入失败测试，验证 Windows 用户目录、互斥量释放和服务就绪超时。
- [ ] 运行 `python -m pytest tests/test_desktop_runtime.py -q`，确认因模块缺失失败。
- [ ] 实现纯函数、Windows 命名互斥量包装和 HTTP 就绪轮询。
- [ ] 重跑定向测试并确认通过。

### Task 2: 桌面启动器与服务生命周期

**Files:**
- Create: `src/excel_splitter/desktop.py`
- Modify: `src/excel_splitter/web/app.py`
- Test: `tests/test_desktop.py`

**Interfaces:**
- Consumes: Task 1 的路径、单实例与就绪轮询接口。
- Produces: `DesktopServer.start() -> str`、`DesktopServer.stop() -> None`、`main() -> int`。

- [ ] 写入失败测试，用可控服务器工厂验证启动、动态端口 URL 和幂等关闭。
- [ ] 运行定向测试并确认失败原因是桌面启动器尚未实现。
- [ ] 使用 Waitress `create_server`、后台线程和 pywebview 窗口实现最小启动流程。
- [ ] 将 Flask 默认输出目录改为可注入且适合冻结环境，不改变测试配置行为。
- [ ] 重跑定向测试并确认通过。

### Task 3: 依赖、构建配置与发布文档

**Files:**
- Modify: `pyproject.toml`
- Modify: `requirements.txt`
- Create: `packaging/excel_splitter.spec`
- Create: `scripts/build_portable.ps1`
- Create: `packaging/使用说明.txt`
- Modify: `README.md`

**Interfaces:**
- Consumes: `excel_splitter.desktop:main` 桌面入口。
- Produces: `dist/Excel拆分工具/Excel拆分工具.exe`。

- [ ] 固定运行与构建依赖版本，并增加 `excel-splitter-desktop` 命令入口。
- [ ] 编写 PyInstaller `onedir`、`windowed` spec，收集模板、静态资源和 pywebview 子模块。
- [ ] 编写可重复构建脚本，清理本项目 `build/dist` 后执行 spec，并复制使用说明。
- [ ] 更新 README 的开发启动、桌面启动、构建与分发步骤。

### Task 4: 构建与验收

**Files:**
- Verify: `dist/Excel拆分工具/**`

**Interfaces:**
- Consumes: Task 3 的构建脚本与 spec。
- Produces: 可内部传输的便携目录及压缩包。

- [ ] 安装固定依赖并运行 `python -m pip check`。
- [ ] 运行 `python -m pytest -q`、`python -m compileall -q src` 和 JavaScript 语法检查。
- [ ] 运行 `scripts/build_portable.ps1`，确认 exe 和 `_internal` 存在。
- [ ] 启动打包 exe，验证本地服务就绪和独立窗口出现。
- [ ] 关闭窗口，确认 exe 进程与监听端口均退出。
- [ ] 生成便携 zip，并记录大小和 SHA-256。

## Self-Review

- 设计范围均映射到四个任务，无自动更新、安装器或业务性能改造。
- 接口名称在任务间一致，无占位需求。
- 项目不是 Git 仓库，因此执行时跳过提交步骤并保留完整验证证据。

