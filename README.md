# Excel工具集

这个项目包含两个高效的Excel处理工具：Excel自动翻译工具和Excel去除重复项工具。

## 工具列表

### 1. Excel自动翻译工具 (`translate_ai.py`)

一个高效的Excel文件翻译工具，支持中英文互译，可以自动识别Excel文件中的文本并进行批量翻译。

#### 功能特点

- 支持多种翻译API（MyMemory、Google、百度翻译、DeepSeek-V3）
- 支持中英文互译
- 一次性翻译多列数据
- 保留原始数据，翻译结果存放在新列
- 自动去重减少API调用次数
- 批量翻译提高效率
- 支持命令行和交互式模式
- CSV中间格式优化，适合处理大文件
- 支持从配置文件、环境变量或.env文件加载API密钥

### 2. Excel去除重复项工具 (`check_duplicates.py`)

一个用于检查Excel文件中指定列重复项的工具，帮助您识别和处理数据中的重复值。

#### 功能特点

- 检查Excel文件中指定列的重复项
- 显示重复项及其所在行号
- 支持交互式和命令行模式
- 可将检查结果保存到文件
- 支持多列依次检查

## 安装

### 系统要求

- Python 3.6+
- Windows/macOS/Linux

### 安装步骤

1. 确保已安装Python
2. 安装所需依赖库：

```bash
# 基本功能库（必需）
pip install openpyxl deep-translator tqdm pandas requests

# 高级功能库（可选，用于配置文件和环境变量支持）
pip install python-dotenv configparser
```

## 使用方法

### Excel自动翻译工具

#### 交互式模式

```bash
python translate_ai.py -i
```

按照提示选择Excel文件、翻译API和翻译方向即可。

#### 命令行模式

```bash
python translate_ai.py -f example.xlsx --zh2en A,B --en2zh C,D --api 1 --batch 20
```

##### 参数说明
- `-f, --file`: Excel文件路径
- `--zh2en`: 需要从中文翻译成英文的列（如A,B,C）
- `--en2zh`: 需要从英文翻译成中文的列（如A,B,C）
- `--api`: 翻译API选择（1=MyMemory, 2=Google, 3=百度, 4=DeepSeek-V3）
- `--batch`: 批量翻译大小（默认10）
- `--use-csv`: 使用CSV中间格式加速翻译（适合大文件）
- `--gen-config`: 生成配置文件模板

#### 使用DeepSeek-V3翻译

DeepSeek-V3是一种高级AI翻译引擎，能提供更高质量的翻译结果。

1. 生成配置文件：
```bash
python translate_ai.py --gen-config
```

2. 编辑`config.ini`文件，填入API密钥。

3. 使用DeepSeek-V3翻译：
```bash
python translate_ai.py -f example.xlsx --zh2en A,B --api 4
```

### Excel去除重复项工具

#### 交互式模式

```bash
python check_duplicates.py -i
```

按照提示选择Excel文件、要检查的列，查看结果。

#### 命令行模式

```bash
python check_duplicates.py example.xlsx -c A
```

##### 参数说明
- 第一个参数: Excel文件路径
- `-c, --column`: 要检查的列（例如：A、B、C等）
- `-i, --interactive`: 使用交互式模式

## 详细文档

更详细的使用说明请参考：
- [Excel自动翻译工具使用文档](excel自动化翻译使用文档.md)
- [Excel去除重复项工具使用文档](excel去除重复项使用文档.md)

## 注意事项

- 百度翻译和DeepSeek-V3 API需要提供相应的密钥和凭证
- 部分地区可能无法使用Google翻译
- 太大的批量处理大小可能导致API调用失败
- 将API密钥等敏感信息保存在环境变量或配置文件中更安全
- 如果不需要配置文件和环境变量功能，无需安装python-dotenv和configparser库 