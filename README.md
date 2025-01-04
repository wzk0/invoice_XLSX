# invoice_XLSX

一个读取当前文件夹(含子文件夹)下所有`png/jpg/pdf格式发票`到xlsx表格的程序，使用百度发票识别API.

## 使用

clone此仓库，安装`openpyxl`，把`发票数据.xlsx`和`app.py`放在发票所在文件夹.

打开`app.py`，填写两个KEY：

```
API_KEY = ""
SECRET_KEY = ""
```

> 程序会读取程序所在文件夹(及子文件夹)内所有`png/jpg/pdf格票文件`. 两个KEY在百度官网自行获取：https://console.bce.baidu.com/ai-engine/ocr/overview/index

随后运行：
```
python3 app.py
```

结果会保存在`发票数据.xlsx`内.

## 注意

`发票数据.xlsx`是输出结果模板，输出结果涵盖19个常用数据.
