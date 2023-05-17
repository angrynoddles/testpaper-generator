# testpaper-generator
组卷工具/试卷自动生成/automatically generate test papers based on excel question bank

参考了[wukai0909](https://github.com/wukai0909)的项目[Generating-WORD-Test-Papers-from-Excel-Question-Bank](https://github.com/wukai0909/Generating-WORD-Test-Papers-from-Excel-Question-Bank)

增加了单元抽取、打乱选项等功能，运行环境python 3.10

## 使用

1. 安装依赖：`pip install -r requirements.txt`
2. 运行： `python create_paper_gui.py`，输出在xlsx题库同目录下output文件夹
3. pyinstaller打包请加选项：`--hidden-import openpyxl.cell._writer`