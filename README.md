# 物品标签生成器

## 功能特性

1. ✅ 读取 `template.docx` 模板文件
2. ✅ 替换模板中的占位符：`物品名字`、`物品编号`、`所在`
3. ✅ 单个物品批量生成标签
4. ✅ 多个物品批量生成标签
5. ✅ 自动生成 Word 文档（.docx）
6. ✅ 根据CSV文件批量生成标签
7. ✅ 导出 JSON 格式记录
8. ✅ 导出 CSV 格式记录

## 环境要求

- Windows 10/11
- .NET 6.0 SDK 或更高版本
- Microsoft Office（用于 PDF 转换，可选）

## 编译步骤

### 1. 安装 .NET SDK

从官网下载并安装：https://dotnet.microsoft.com/download/dotnet/6.0

### 2. 创建项目目录结构

```
LabelGenerator/
├── Program.cs
├── MainForm.cs
├── LabelGenerator.csproj
├── Label_Generator.sln
└── template.docx
```

### 3. 编译项目

打开命令提示符，进入项目目录，执行：

```bash
dotnet restore
dotnet build
```

### 4. 发布为单个 EXE 文件

```bash
dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true /p:IncludeNativeLibrariesForSelfExtract=true
```

编译后的 exe 文件位于：
```
bin\Release\net6.0-windows\win-x64\publish\LabelGenerator.exe
```

## 使用说明

### 准备模板文件

在程序目录下放置 `template.docx` 文件，模板中包含以下占位符：
- `物品名字` - 将被替换为实际物品名称
- `物品编号` - 将被替换为物品编号
- `所在` - 将被替换为所在室信息

### 单个物品标签生成

1. 打开程序，进入"单个物品"标签页
2. 输入以下信息：
   - 物品名字：例如 "笔记本电脑"
   - 物品数量：例如 10
   - 起始编号：例如 1001
   - 所在室：例如 "301室"
3. 点击"生成标签"按钮

### 批量物品标签生成

1. 打开程序，进入"批量物品"标签页
2. 在表格中输入多行物品信息：
   ```
   物品名字      | 数量 | 起始编号 | 所在室
   笔记本电脑    | 5    | 1001    | 212
   显示器        | 10   | 2001    | 210
   鼠标          | 20   | 3001    | 208
   ```
3. 点击"生成批量标签"按钮

### 输出文件

程序会在同目录下生成以下文件：

- `标签_20250101_120000.docx` - Word 文档
- `标签_20250101_120000.pdf` - PDF 文档
- `标签_20250101_120000.json` - JSON 格式记录
- `标签_20250101_120000.csv` - CSV 格式记录

### JSON 格式示例

```json
[
  {
    "Name": "笔记本电脑",
    "Number": "1001",
    "Location": "301室"
  },
  {
    "Name": "笔记本电脑",
    "Number": "1002",
    "Location": "301室"
  }
]
```

### CSV 格式示例

```csv
物品名字,物品编号,所在室
"笔记本电脑","1001","301室"
"笔记本电脑","1002","301室"
```

## 注意事项

1. **模板格式保持**：程序会完整复制模板的格式、字体、样式、纸张大小等
3. **分页**：每个物品标签占一页，自动分页
4. **编号递增**：物品编号从起始编号开始自动递增

## 故障排除

### 模板文件未找到

**问题**：提示 "未找到 template.docx 文件"

**解决方案**：
1. 确保 `template.docx` 文件与 exe 在同一目录
2. 检查文件名是否正确（注意大小写）

### 编译错误

**问题**：编译时提示缺少依赖

**解决方案**：
```bash
dotnet restore
dotnet clean
dotnet build
```

## 技术栈

- **框架**：.NET 6.0 WinForms
- **Word 处理**：DocX (Xceed.Words.NET)
- **数据格式**：System.Text.Json, CSV

## 许可证

本程序为内部使用工具，请遵守MIT开源库的许可证。