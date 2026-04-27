# ReportFlow 产品需求文档

## 1. 产品定位

ReportFlow 是一个面向非技术用户的 Excel/WPS 表格报表规则录制工具。产品不再尝试复刻表格界面，而是直接调用用户电脑上安装的 Microsoft Excel 或 WPS 表格。

用户在真正的 Excel 中完成操作，ReportFlow 在旁边捕获操作结果并生成可复用规则。后续用户上传或打开同类文件时，可以复用这些规则生成结果报表。

## 2. 核心原则

```text
1. 不仿 Excel，直接用 Excel。
2. 不让用户配置复杂表单，让用户按原来的 Excel 习惯操作。
3. ReportFlow 负责录制、解释、复用和批处理。
4. 结果必须可导出为方案 JSON，并可再次执行。
```

## 3. 目标用户

```text
1. 行政人员
2. 计划人员
3. 售后人员
4. 运营助理
5. 部门文员
6. 统计报表人员
```

## 4. 核心流程

### 创建规则

```text
打开 ReportFlow
↓
点击“打开 Excel 并开始录制”
↓
ReportFlow 启动 Microsoft Excel 或 WPS 表格并打开工作簿
↓
用户在 Excel 中正常处理文件
↓
用户回到 ReportFlow 点击“捕获当前操作为规则”
↓
ReportFlow 根据打开时快照、当前快照和 Excel 状态生成规则
↓
用户导出方案 JSON
```

### 复用规则

```text
打开同类 Excel
↓
导入方案 JSON
↓
点击执行
↓
生成结果 Excel
```

## 5. 规则捕获范围

MVP 捕获：

```text
1. 删除列
2. 隐藏列，按删除列处理
3. 重命名列
4. 新增空列/固定值列
5. 保留/调整列
6. 单元格修改
7. 新增或修改公式列
8. 自动筛选条件
9. 排序字段
10. 列宽、行高、数字格式、字体加粗、字体颜色、填充色
11. 新增图表
12. Sheet 新增/删除
13. 跨 Sheet 公式引用
```

后续增强：

```text
1. 更完整的宏动作事件捕获
2. 复制粘贴区域识别
3. 数据透视表识别
4. 条件格式和复杂样式模板复用
5. VBA 宏导入和解析
6. AI 将用户自然语言需求转换为 Excel 公式
7. 更精细的图表类型、图例、坐标轴和数据源还原
```

## 6. 技术方案

```text
1. Python 3.11
2. Tkinter 控制台
3. pywin32 调用 Microsoft Excel/WPS 表格 COM
4. openpyxl 执行规则并生成结果文件
5. PyInstaller 打包 exe
```

## 7. 依赖条件

```text
1. Windows
2. 本机安装 Microsoft Excel 或 WPS 表格
3. Python 3.11 开发/打包环境
```

## 8. 验收标准

```text
1. ReportFlow 可以启动 Microsoft Excel 或 WPS 表格。
2. 用户可以在真正的表格软件中操作。
3. ReportFlow 可以捕获当前工作簿状态并生成规则。
4. 规则列表可读。
5. 方案可导出 JSON。
6. 方案可导入复用。
7. 可以执行方案并生成结果 Excel。
```

## 9. 不做

```text
1. 不再自研表格 UI。
2. 不再使用前后端网页。
3. 不再要求 MySQL。
4. 不完整替代 Excel。
```
