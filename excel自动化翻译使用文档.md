# Excel自动翻译工具使用说明

## 目录
1. [安装Python](#安装python)
2. [安装必要的库](#安装必要的库)
3. [下载翻译工具](#下载翻译工具)
4. [使用翻译工具](#使用翻译工具)
   - [交互式模式](#交互式模式)
   - [命令行模式](#命令行模式)
5. [使用DeepSeek-V3翻译](#使用deepseek-v3翻译)
   - [配置DeepSeek-V3 API](#配置deepseek-v3-api)
   - [使用配置文件](#使用配置文件)
   - [使用环境变量](#使用环境变量)
   - [使用.env文件](#使用env文件)
6. [优化大文件翻译](#优化大文件翻译)
7. [常见问题解答](#常见问题解答)

## 安装Python

### 第1步：下载Python安装程序

1. 打开浏览器，访问Python官网：https://www.python.org/downloads/
2. 点击黄色的"Download Python 3.x.x"按钮（x.x表示最新版本号）
3. 安装文件会自动下载到您的电脑上

### 第2步：安装Python

1. 找到并双击下载的安装文件（通常在"下载"文件夹中）
2. 在安装界面上，**务必勾选**"Add Python to PATH"选项（这非常重要！）
3. 点击"Install Now"（现在安装）按钮
4. 等待安装完成
5. 看到"Setup was successful"（安装成功）的提示后，点击"Close"（关闭）按钮

### 第3步：验证安装

1. 按下`Win + R`键，打开"运行"对话框
2. 输入`cmd`并点击"确定"，打开命令提示符窗口
3. 在命令提示符窗口中输入：`python --version`并按回车键
4. 如果看到类似`Python 3.x.x`的输出，说明安装成功

## 安装必要的库

Excel翻译工具需要安装几个特定的Python库才能正常工作。请按照以下步骤安装：

### 基本库（必需）

1. 按下`Win + R`键，打开"运行"对话框
2. 输入`cmd`并点击"确定"，打开命令提示符窗口
3. 在命令提示符窗口中输入以下命令并按回车键：
   ```
   pip install openpyxl deep-translator tqdm pandas requests
   ```
4. 等待安装完成，会看到"Successfully installed"（安装成功）的提示

### 高级功能库（可选）

如果您需要使用配置文件或.env文件功能来管理API密钥，可以安装以下库：
```
pip install python-dotenv configparser
```

这些库不是必需的，如果没有安装，程序仍然可以通过命令行参数或手动输入方式使用。

## 下载翻译工具

1. 将`test_optimized.py`文件保存到您的电脑上（可以保存到桌面或任何您方便找到的位置）

## 使用翻译工具

### 交互式模式

交互式模式是最简单的使用方式，程序会引导您完成所有步骤。

1. 将您要翻译的Excel文件放在与`test_optimized.py`相同的文件夹中
2. 按下`Win + R`键，打开"运行"对话框
3. 输入`cmd`并点击"确定"，打开命令提示符窗口
4. 使用`cd`命令导航到`test_optimized.py`所在的文件夹，例如：
   ```
   cd C:\Users\您的用户名\Desktop
   ```
5. 输入以下命令并按回车键启动交互式模式：
   ```
   python test_optimized.py -i
   ```
6. 按照程序提示进行操作：
   - 选择要翻译的Excel文件
   - 选择翻译API（推荐选择1: MyMemory翻译，或者4: DeepSeek-V3提供更高质量翻译）
   - 选择要翻译的列（从中文翻译成英文和/或从英文翻译成中文）
   - 设置批量翻译大小（可以直接按回车使用默认值）
   - 选择是否使用CSV中间格式（对于大文件推荐使用）
   - 等待翻译完成
   - 选择是否打开生成的Excel文件

### 命令行模式

如果您已经熟悉该工具，可以使用命令行模式更快速地完成翻译：

1. 按下`Win + R`键，打开"运行"对话框
2. 输入`cmd`并点击"确定"，打开命令提示符窗口
3. 使用`cd`命令导航到`test_optimized.py`所在的文件夹
4. 输入以下命令并按回车键：
   ```
   python test_optimized.py -f 您的Excel文件名.xlsx --zh2en A,B --en2zh C,D
   ```
   其中：
   - `您的Excel文件名.xlsx`是您要翻译的Excel文件名
   - `--zh2en A,B`表示将A列和B列从中文翻译成英文
   - `--en2zh C,D`表示将C列和D列从英文翻译成中文

## 使用DeepSeek-V3翻译

DeepSeek-V3是一款强大的AI翻译引擎，可以提供更高质量的翻译结果。使用DeepSeek-V3需要API密钥。

### 配置DeepSeek-V3 API

有三种方式配置DeepSeek-V3的API密钥和URL：

1. 使用配置文件（推荐，需要安装configparser库）
2. 使用环境变量
3. 使用.env文件（需要安装python-dotenv库）

选择一种最适合您的方式进行配置。

### 使用配置文件

1. 确保已安装configparser库：
   ```
   pip install configparser
   ```

2. 生成配置文件模板：
   ```
   python test_optimized.py --gen-config
   ```

3. 这将在当前目录创建`config.ini`文件，用文本编辑器打开并编辑：
   ```ini
   [api]
   deepseek_key = YOUR_DEEPSEEK_API_KEY
   deepseek_url = https://api.deepseek.com/v1/chat/completions
   baidu_appid = YOUR_BAIDU_API_ID
   baidu_key = YOUR_BAIDU_API_KEY
   ```

4. 将`YOUR_DEEPSEEK_API_KEY`替换为您的实际API密钥。

配置文件会自动从以下路径查找：
- 当前目录的`config.ini`
- 用户主目录下的`~/.excel_translator/config.ini`
- 脚本所在目录的`config.ini`

### 使用环境变量

您可以设置以下环境变量：

#### Windows系统：

1. 按下`Win + R`键，输入`sysdm.cpl`，点击"确定"
2. 在系统属性窗口，切换到"高级"选项卡
3. 点击"环境变量"按钮
4. 在"用户变量"部分，点击"新建"按钮
5. 添加以下变量：
   - 变量名：`DEEPSEEK_API_KEY`，变量值：您的API密钥
   - 变量名：`DEEPSEEK_API_URL`，变量值：API URL（可选）

#### Mac/Linux系统：

在终端中执行：
```bash
export DEEPSEEK_API_KEY=your_api_key
export DEEPSEEK_API_URL=your_api_url
```

要永久保存，可以将这些命令添加到您的`.bashrc`或`.zshrc`文件中。

### 使用.env文件

1. 确保安装了python-dotenv库：
   ```
   pip install python-dotenv
   ```

2. 在工作目录创建一个名为`.env`的文件
3. 添加以下内容：
   ```
   DEEPSEEK_API_KEY=your_api_key
   DEEPSEEK_API_URL=your_api_url
   ```
4. 保存文件

程序会自动读取这个文件中的环境变量。

## 优化大文件翻译

对于大型Excel文件（超过50MB或包含大量数据），可以使用CSV中间格式来提高翻译效率：

1. 在交互式模式中，当询问"是否使用CSV中间格式加速翻译"时选择"y"
2. 在命令行模式中，添加`--use-csv`参数：
   ```
   python test_optimized.py -f 大文件.xlsx --zh2en A,B --use-csv
   ```

CSV转换的优势：
- 降低内存使用量，避免处理大文件时内存不足
- 流式处理数据，而不是一次加载整个文件
- 对唯一文本进行翻译，减少重复翻译
- 最终结果仍然是Excel格式，不影响使用

当使用DeepSeek-V3翻译大量文本时，CSV模式尤其有效，因为它可以更好地利用DeepSeek-V3处理长文本的能力。

## 常见问题解答

### 问题1：运行程序时提示"python不是内部或外部命令"

**解决方法**：重新安装Python，确保在安装时勾选"Add Python to PATH"选项。

### 问题2：程序提示找不到某个模块

**解决方法**：根据错误信息安装缺失的库：

- **基本功能库**（翻译功能必需）：
  ```
  pip install deep-translator
  pip install openpyxl
  pip install tqdm
  pip install pandas
  pip install requests
  ```

- **高级功能库**（配置文件和环境变量功能）：
  ```
  pip install python-dotenv
  pip install configparser
  ```

### 问题3：提示"No module named 'dotenv'"或"No module named 'configparser'"

**解决方法**：这些是可选库，用于支持配置文件和.env文件功能。如果您需要这些功能，请安装：
```
pip install python-dotenv configparser
```
如果不安装，您仍然可以通过命令行参数或交互式界面手动输入API密钥。

### 问题4：翻译过程中出现错误或翻译结果不正确

**解决方法**：
- 确保您的网络连接正常
- 尝试减小批量翻译大小（建议值为5-10）
- 尝试使用不同的翻译API（例如从Google切换到MyMemory或DeepSeek-V3）
- 如果使用DeepSeek-V3，确认API密钥正确

### 问题5：如何知道我的Excel文件中哪些列需要翻译？

**解决方法**：打开Excel文件，查看哪些列包含中文内容（需要翻译成英文）和哪些列包含英文内容（需要翻译成中文）。在程序中，列是按字母表示的：A代表第一列，B代表第二列，以此类推。

### 问题6：翻译后的结果在哪里？

**解决方法**：翻译后的结果会保存在与原Excel文件相同的文件夹中，文件名格式为"原文件名_translated.xlsx"。翻译结果会在原始列的右侧新增列显示。

### 问题7：DeepSeek-V3比其他翻译API好在哪里？

**解决方法**：DeepSeek-V3是基于大型语言模型的翻译引擎，相比传统翻译API有以下优势：
- 更好的上下文理解能力，翻译结果更自然
- 可以一次处理更多文本，提高翻译效率
- 对专业术语和特定领域内容翻译准确性更高

### 问题8：如何更改DeepSeek-V3的批量翻译大小？

**解决方法**：代码中默认设置了3500字符的批量大小，这是根据DeepSeek-V3的能力设定的平衡值。如果需要修改，可以编辑代码中的`max_chars`变量（在`batch_translate`函数中）。 