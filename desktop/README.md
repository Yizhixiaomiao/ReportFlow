# ReportFlow Desktop

这是 ReportFlow 的 Excel 原生伴侣版本。它不再自己绘制表格，而是直接调用 Microsoft Excel。

## 准备

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
.\build_exe.ps1 -InstallDeps
```

## 运行

```powershell
python reportflow_desktop.py
```

## 打包

```powershell
.\build_exe.ps1
```

输出：

```text
dist/ReportFlowDesktop/ReportFlowDesktop.exe
```

## 用法

1. 左侧点击“打开 Excel 并开始录制”。
2. 在真正的 Excel 里操作。
3. 点击“捕获当前操作为规则”。
4. 左侧查看已生成规则。
5. 在规则列表下方使用“函数查询 / 生成”。
6. 左下角“设置 -> 操作文档”查看完整说明。
7. 导出方案，或直接执行生成结果。

当前通过快照差异和 Excel 当前状态捕获规则，包括列删除、单元格修改、公式列、筛选和排序。
