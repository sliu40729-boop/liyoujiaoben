# 学生名单比对

用 Excel 维护自己负责的学生名单。可以按姓名或学号查询，导入 XLSX、XLS 或 CSV 批量比对，也可以识别照片后校对并比对。结果可下载为包含已匹配、待核对、未匹配三个工作表的 Excel。

## 启动

安装 Python 3.10 或更新版本，在项目目录执行：

```sh
python -m pip install -r requirements.txt
python -m streamlit run app.py --server.address 127.0.0.1
```

在浏览器中打开启动命令显示的地址。首次使用，在左侧上传包含「学号」「姓名」的底库并点击「保存为默认底库」，程序会在项目目录保存 `底库.xlsx`。也可手动将底库另存为这个文件名。云工作区已有的《群成员名单（学号与姓名）》会自动加载。启动后可以按需更新底库。

## 照片识别

安装 Tesseract 和 `chi_sim`、`eng` 语言包，并使用 `tesseract --list-langs` 检查。Debian 或 Ubuntu 可执行：

```sh
sudo apt install tesseract-ocr tesseract-ocr-chi-sim tesseract-ocr-eng
```

macOS 可用 Homebrew 安装 `tesseract` 和 `tesseract-lang`。Windows 可使用 Tesseract 文档列出的 [Windows 安装程序](https://tesseract-ocr.github.io/tessdoc/Installation.html)，安装时选简体中文和英文。若程序不在 PATH 中，可将环境变量 `TESSERACT_CMD` 指向可执行文件。图片由运行本工具的机器处理。识别后检查每行是否对应一名学生，再点击「确认并比对」。

## 匹配与核对

- 学号优先。输入记录同时带有姓名时，学号命中但姓名冲突的记录进入「待核对」。
- 只有姓名时按完整姓名匹配。底库有同名时，列出候选供核对。
- 底库来自群昵称。原表中学号缺失或原数字部分为四位的两条记录，请按原表「待核对」页核实后更新。
