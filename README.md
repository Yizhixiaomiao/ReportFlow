# ReportFlow

ReportFlow 现在改为 **Excel 原生伴侣模式**：不再自己仿一个表格，也不再让用户在笨重的表单里选操作。应用会直接打开 Microsoft Excel，用户就在真正的 Excel 里筛选、排序、删列、写公式、改数据；ReportFlow 只负责在旁边录制这些变化，生成可复用规则。

## 运行方式

首次准备环境：

```powershell
cd D:\Workspace\ReportFlow\desktop
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
.\build_exe.ps1 -InstallDeps
```

运行 exe：

```text
D:\Workspace\ReportFlow\desktop\dist\ReportFlowDesktop\ReportFlowDesktop.exe
```

开发模式：

```powershell
cd D:\Workspace\ReportFlow\desktop
.\.venv\Scripts\Activate.ps1
python reportflow_desktop.py
```

## 使用流程

左侧是主要操作区：

1. 点击“打开 Excel 并开始录制”。
2. 在真正的 Excel 里正常操作：筛选、排序、删除列、修改单元格、新增公式列。
3. 回到 ReportFlow，点击“捕获当前操作为规则”。
4. 在左侧“已生成规则”中检查规则。
5. 在规则列表下方使用“函数查询 / 生成”，可复制公式或写入 Excel 当前单元格。
6. 点击“一键加载规则”复用 JSON 方案，或点击“导出规则”保存当前规则。
7. 点击“执行并生成结果”输出结果 Excel。

左下角“设置 -> 操作文档”内置完整使用说明。

## 当前可录制内容

- 删除列、保留/调整列。
- 单元格修改。
- 新增 Excel 公式列。
- Excel 自动筛选条件。
- Excel 排序字段。

## 依赖

- Windows
- Microsoft Excel
- Python 3.11
- openpyxl
- pywin32
- PyInstaller

## 注意

这不是把 Excel 重新实现一遍，而是调用用户电脑上已经安装的 Excel。这样体验最接近用户真实习惯，也避免了假表格 UI 操作麻烦的问题。
