# Invoice-merger

一个用于合并发票和行程单的工具，借助 ChatGPT 完成。

## 介绍

Invoice Merger 是一个用于合并发票和行程单的工具。它能够从指定的文件夹中提取发票和行程单，并将它们合并到一个 Word 文档中。工具提供了一个简单易用的界面，让用户能够方便地选择输入文件夹和输出路径，并在文档开头显示发票的总金额。

## 安装依赖

在运行 Invoice Merger 之前，请确保已经安装了以下依赖项：

- Python 3.x：`https://www.python.org/downloads/`
- PyPDF2 库：通过运行 `pip install PyPDF2` 进行安装
- docx 库：通过运行 `pip install python-docx` 进行安装
- pdf2image 库：通过运行 `pip install pdf2image` 进行安装

## 使用方法

### 使用代码

1. 克隆或下载本仓库到本地。
1. 打开命令行终端，并进入项目目录。
1. 安装所需的 Python 依赖（参考上述的安装依赖步骤）。
1. 运行 `python invoice_merger.py` 命令启动工具。
1. 在工具界面中选择输入文件夹和输出路径。
1. 点击运行按钮开始合并操作。
1. 合并完成后，在指定的输出路径将生成一个包含合并后发票和行程单的 Word 文档。

### 使用软件

1. 下载软件到本地。
2. 双击打开软件。
3. 在工具界面中选择输入文件夹和输出路径。
4. 点击运行按钮开始合并操作。
5. 合并完成后，在指定的输出路径将生成一个包含合并后发票和行程单的 Word 文档。

## 注意事项

- 输入文件夹应仅包含发票和行程单的 PDF 文件。
- 发票文件的命名格式应符合工具要求，以便正确提取发票金额。
- 合并后的发票总金额将显示在生成的 Word 文档的结尾。

## 示例截图

![image](https://github.com/Shitao5/Invoice-merger/assets/68451957/9835dcbf-b138-4ba7-90b4-6f7e499f7ef2)

## 贡献

如果你对该项目有任何改进或建议，请随时提出 Issue 或提交 Pull Request。
