# Excel 通用拆表工具

本地运行的 `.xlsx` 多 Sheet 拆分工具。每个 Sheet 可以独立选择表头行、直接拆分、完整保留或通过关联键匹配，并可输出公式版、结果值版或两者。

## 开发环境

- Windows 10/11 x64
- Python 3.11
- Flask 3.1.3
- openpyxl 3.1.5
- pywebview 6.2.1
- Waitress 3.0.2

安装项目与构建依赖：

```powershell
D:\conda_envs\edc\python.exe -m pip install -e ".[build]"
```

## 启动 Web 开发版

```powershell
cd D:\EDC\hr_work\excel_splitter_tool
.\scripts\start_web.ps1
```

浏览器访问 `http://127.0.0.1:5000/`。

## 启动桌面开发版

```powershell
D:\conda_envs\edc\python.exe -m excel_splitter.desktop
```

桌面版使用动态本地端口，并在独立 WebView2 窗口中显示现有界面。关闭窗口后本地服务自动停止。

## 构建便携版

```powershell
.\scripts\build_portable.ps1
```

构建产物：

```text
dist\Excel拆分工具\Excel拆分工具.exe
dist\ExcelSplitter-portable.zip
```

内部传输时发送完整 zip。用户必须先完整解压，再双击主程序；目标电脑不需要 Python，但需要 Microsoft Edge WebView2 Runtime。

## 验证

```powershell
.\scripts\verify.ps1
```

或直接运行：

```powershell
D:\conda_envs\edc\python.exe -m pytest -q
node --check src\excel_splitter\web\static\app.js
D:\conda_envs\edc\python.exe -m compileall -q src
```

## 数据目录

- 桌面版默认输出：`%USERPROFILE%\Documents\Excel拆分工具输出`
- 日志：`%LOCALAPPDATA%\ExcelSplitter\logs\excel-splitter.log`
- 上传文件：系统临时目录，正常退出时清理

程序只监听 `127.0.0.1`，不会将工作簿上传到互联网。
