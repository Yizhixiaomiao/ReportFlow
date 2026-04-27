# ReportFlow

ReportFlow 现在改为 **表格原生伴侣模式**：不再自己仿一个表格，也不再让用户在笨重的表单里选操作。应用会直接打开 Microsoft Excel 或 WPS 表格，用户就在真正的表格软件里筛选、排序、删列、写公式、改数据；ReportFlow 只负责在旁边录制这些变化，生成可复用规则。

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
2. 在真正的 Excel/WPS 表格里正常操作：筛选、排序、删除/隐藏列、重命名列、修改单元格、新增固定值列、新增或修改公式列、调整格式、制作图表、跨 Sheet 写公式。
3. 回到 ReportFlow，点击“捕获当前操作为规则”。
4. 在左侧“已生成规则”中检查规则。
5. 在规则列表下方使用“函数查询 / 生成”，可复制公式或写入 Excel 当前单元格。
6. 点击“一键加载规则”复用 JSON 方案，或点击“导出规则”保存当前规则。
7. 点击“执行并生成结果”输出结果 Excel。

当前会录制数据操作、基础格式变化、图表新增和跨 Sheet 公式引用。复杂图表样式、数据透视表、条件格式会先作为后续增强项。

左下角“设置 -> 操作文档”内置完整使用说明。

设置中可以配置：

- 表格内核偏好：自动、优先 Excel、优先 WPS。
- 捕获范围：数据操作、格式修改、图表制作、Sheet 新增/删除、跨 Sheet 公式。
- 最大捕获行数。
- 生成后是否自动打开结果。
- WPS `.et` 临时转换文件是否保留。

## 当前可录制内容

- 删除列、保留/调整列。
- 单元格修改。
- 新增 Excel 公式列。
- Excel 自动筛选条件。
- Excel 排序字段。

## 依赖

- Windows
- Microsoft Excel 或 WPS 表格
- Python 3.11
- openpyxl
- pywin32
- PyInstaller

## 注意

这不是把 Excel/WPS 重新实现一遍，而是调用用户电脑上已经安装的表格软件。这样体验最接近用户真实习惯，也避免了假表格 UI 操作麻烦的问题。
