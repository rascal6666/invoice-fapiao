# 发票识别器 (Invoice Recognizer)

这是一个基于AI的发票识别工具，可以从PDF发票中提取结构化数据并生成Excel表格。

## 功能特点

- 🔍 **智能识别**：基于DeepSeek AI的发票信息提取
- 📊 **批量处理**：支持同时处理多个PDF文件
- 💾 **缓存机制**：避免重复调用AI，节省费用
- 🛡️ **容错处理**：单个文件失败不影响整体处理
- 📱 **友好界面**：直观的GUI操作界面
- 🔐 **安全配置**：加密存储API密钥，避免明文泄露
- 🔑 **申请指南**：内置API密钥申请步骤指导
- ⏱️ **时间戳**：文件名包含生成时间，便于管理
- 📈 **实时进度**：显示处理进度和当前状态

## 安装要求

### 系统要求
- Windows 10/11
- Python 3.8+
- 网络连接

### 依赖包
```
tkinter
openpyxl
pdfplumber
requests
```

### API密钥
- 需要DeepSeek API密钥
- 格式：`sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx`
- 可在程序界面配置和保存

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 1. GUI界面（推荐）

```bash
# 运行GUI程序
python gui_app.py

# 或者双击运行批处理文件
run_gui.bat
```

GUI界面功能：
- 📋 **使用说明**: 详细的功能说明和注意事项
- 📁 **目录选择**: 选择包含PDF文件的目录
- 📊 **进度显示**: 实时显示处理进度和当前文件
- 📝 **日志记录**: 显示详细的处理日志
- 🎯 **一键处理**: 点击按钮即可开始批量处理

### 2. 命令行界面

```python
from entry import parse_invoice_from_pdf

# 解析单个PDF文件
invoice_info = parse_invoice_from_pdf("path/to/invoice.pdf")

# 查看解析结果
print(f"发票号码: {invoice_info.invoice_number}")
print(f"销方名称: {invoice_info.seller_name}")
print(f"购方名称: {invoice_info.buyer_name}")
```

### 3. 批量处理

```python
from entry import process_directory_to_xlsx

# 处理目录中所有PDF文件并生成Excel表格
process_directory_to_xlsx("./pdf_files", "发票数据汇总.xlsx")
```

### 4. 使用示例脚本

```bash
# 创建PDF文件目录
mkdir pdf_files

# 将PDF文件放入目录
# ...

# 运行批量处理
python example_usage.py
```

### 5. 测试功能

```bash
# 测试缓存机制
python test_cache.py

# 测试错误处理功能
python test_error_handling.py
```

## 打包成可执行程序

### 自动打包

```bash
# 运行打包脚本
python build_exe.py
```

打包脚本会自动：
1. 安装PyInstaller
2. 创建配置文件
3. 构建可执行文件
4. 生成`dist/发票识别器.exe`

### 手动打包

```bash
# 安装PyInstaller
pip install pyinstaller

# 创建图标（可选）
python create_icon.py

# 打包程序
pyinstaller --onefile --windowed --icon=icon.ico --name="发票识别器" gui_app.py
```

### 打包后的文件结构

```
dist/
└── 发票识别器.exe    # 可执行文件
```

### 输出文件说明

程序会生成以下文件：
- **Excel文件**：`发票数据汇总_YYYYMMDD_HHMMSS.xlsx`（包含时间戳）
- **缓存文件**：`cache_res_原文件名.json`（避免重复调用AI）

## 缓存机制

为了提高效率并节省API调用费用，系统实现了智能缓存机制：

### 缓存文件命名规则
- 缓存文件格式：`cache_res_原文件名.json`
- 位置：与PDF文件在同一目录

### 缓存逻辑
1. **首次解析**: 调用AI解析PDF，生成缓存文件
2. **重复解析**: 如果缓存文件存在，直接读取缓存，跳过AI调用
3. **缓存失效**: 如果缓存文件损坏，自动重新解析

### 缓存优势
- ⚡ **速度提升**: 缓存读取比AI调用快10-100倍
- 💰 **成本节省**: 避免重复的API调用费用
- 🔄 **结果一致**: 确保相同文件解析结果一致

## 容错处理

系统具备完善的错误处理机制：

### 错误处理策略
1. **单个文件失败**: 不影响其他文件的处理
2. **错误信息记录**: 在Excel的"备注"列显示详细错误信息，包含文件名称
3. **错误行标识**: 错误行使用红色背景标识
4. **继续处理**: 即使部分文件失败，仍会生成包含成功解析数据的Excel
5. **文件定位**: 所有错误信息都包含具体的文件名称，便于快速定位问题文件

### 错误类型
- PDF文件损坏或无法读取
- AI解析失败或返回异常数据
- 网络连接问题
- 文件格式不支持

## Excel表格字段说明

生成的Excel表格包含以下28个字段：

| 序号 | 字段名 | 说明 |
|------|--------|------|
| 1 | 序号 | 自动生成的序号 |
| 2 | 发票代码 | 发票代码 |
| 3 | 发票号码 | 发票号码（留空） |
| 4 | 数电发票号码 | 数电发票号码 |
| 5 | 销方识别号 | 销售方纳税人识别号 |
| 6 | 销方名称 | 销售方名称 |
| 7 | 购方识别号 | 购买方纳税人识别号 |
| 8 | 购买方名称 | 购买方名称 |
| 9 | 开票日期 | 发票开票日期 |
| 10 | 税收分类编码 | 税收分类编码 |
| 11 | 特定业务类型 | 特定业务类型 |
| 12 | 货物或应税劳务名称 | 货物或应税劳务名称 |
| 13 | 名称 | 货物名称 |
| 14 | 规格型号 | 货物规格型号 |
| 15 | 单位 | 计量单位 |
| 16 | 数量 | 货物数量 |
| 17 | 单价 | 单价 |
| 18 | 金额 | 金额 |
| 19 | 税率 | 税率 |
| 20 | 税额 | 税额 |
| 21 | 价税合计 | 价税合计 |
| 22 | 发票来源 | 发票来源 |
| 23 | 发票票种 | 发票类型 |
| 24 | 发票状态 | 发票状态 |
| 25 | 是否正数发票 | 是否为正数发票 |
| 26 | 发票风险等级 | 发票风险等级 |
| 27 | 开票人 | 开票人 |
| 28 | 备注 | 备注信息或错误信息 |

## 注意事项

1. **文件格式**: 仅支持PDF格式的发票文件
2. **AI识别**: 识别准确性依赖于AI模型，复杂格式的发票可能需要人工校验
3. **批量处理**: 每个PDF文件可能包含多个货物项目，会生成多行数据
4. **错误处理**: 单个文件处理失败不会影响其他文件的处理
5. **缓存管理**: 缓存文件会占用磁盘空间，可手动删除不需要的缓存文件
6. **网络要求**: 首次解析需要网络连接调用AI API
7. **系统要求**: GUI程序需要Windows 7或更高版本

## 配置说明

在 `entry.py` 中配置您的DeepSeek API密钥：

```python
DEEP_SEEK_KEY = 'your-api-key-here'
DEEP_SEEK_API_HOST = 'https://api.deepseek.com'
```

## 项目结构

```
项目目录/
├── gui_app.py              # GUI主程序
├── entry.py                # 核心处理逻辑
├── requirements.txt        # 依赖包列表
├── .gitignore             # Git忽略文件配置
├── README.md              # 项目说明文档
├── 使用指南.md            # 详细使用指南
├── build_exe.py           # 可执行文件构建脚本
├── create_icon.py         # 图标生成脚本
├── run_gui.bat            # Windows启动脚本
└── dist/                  # 打包输出目录
    └── 发票识别器.exe     # 可执行文件
```

## 版本控制

### Git配置
项目包含完整的`.gitignore`文件，自动忽略以下文件：
- **Python缓存文件**：`__pycache__/`、`*.pyc`
- **虚拟环境**：`venv/`、`.env`
- **IDE配置**：`.idea/`、`.vscode/`
- **系统文件**：`.DS_Store`、`Thumbs.db`
- **项目特定文件**：
  - API密钥配置：`~/.invoice_recognizer/`
  - 缓存文件：`cache_res_*.json`
  - 生成的Excel文件：`*.xlsx`
  - 测试文件：`test_*.py`、`demo_*.py`

## 许可证

本项目仅供学习和研究使用。 