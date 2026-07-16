# Excel 通用拆表工具

本地运行的 `.xlsx` 多 Sheet 拆分工具。纯本地运行，不上传任何数据。

## 它能做什么

把一个包含多个 Sheet 的 Excel 工作簿，按某个字段（如部门、团队）拆分成多个独立的文件。

**举个例子**：有一张全员工时表，里面有"人员"、"项目"两个 Sheet，按"部门"字段拆分成 "临床部.xlsx"、"研发部.xlsx"……每个拆分结果里都保留原工作簿中该部门的全部 Sheet 数据。

每个 Sheet 可以单独配置处理方式：
- **直接拆分** — 按本 Sheet 的某个字段拆分（如按部门拆分）
- **完整保留** — 不拆分，原样复制到每个输出文件中
- **基准表引用** — 以某个 Sheet 的拆分为基准，其它 Sheet 通过关联键匹配归属
- **关联键匹配** — 按基准表中定义的分组关系，将本 Sheet 的数据归入对应分组

输出版本可选 **公式版**（保留原始 Excel 公式）、**结果值版**（公式已转为计算结果值），或两者同时输出。

## v0.2 更新内容

- **支持加密文件** — 打开加密的 xlsx 文件时自动弹出密码输入框，解锁后正常使用
- **支持输出加密** — 输出设置里可勾选加密，给每个拆分结果文件设置密码保护
- **大幅提升速度** — 同一会话内不再反复读取源文件，大数据量时预览和加载明显变快
- **默认输出到源文件目录** — 拆好的文件会保存在源文件旁边，自动创建 `{文件名}_拆分结果` 文件夹
- **打开文件更方便** — 拆分完成后可以直接点击"打开文件"或"打开所在文件夹"
- **全选 Sheet** — 第二步新增全选按钮，不用一个个勾选
- **默认保留整表** — 第三步新增 Sheet 时默认"完整保留"，减少误拆分
- **去掉多余的选项** — 拆分值选择不再有"全部值/手动勾选"切换，直接勾选要拆分的值即可，更直观
- **修复密码弹窗问题** — 之前打开页面会无故弹出密码窗口，现已修复

## 开发环境

- Windows 10/11 x64
- Python 3.11
- Flask 3.1.3
- openpyxl 3.1.5
- pywebview 6.2.1
- Waitress 3.0.2
- msoffcrypto-tool 6.0.0+

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

## 输出目录

拆分完成后文件默认输出到源文件所在目录下的 `{文件名}_拆分结果` 子文件夹。Web 开发版可手动设置输出路径；桌面版直接输出到源文件同目录下。

## 数据与安全

- 上传文件：系统临时目录，正常退出时清理
- 日志：`%LOCALAPPDATA%\ExcelSplitter\logs\excel-splitter.log`
- 程序只监听 `127.0.0.1`，不会将工作簿上传到互联网
