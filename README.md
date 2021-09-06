一个python脚本，用于自动把各种wps支持的文档，借助windows下的wps自动把各种文档转换为pdf文档，

#### 安装
1. 安装最新版的 python3 ，安装时记得把 python 放到系统 Path 里面
2. 使用 pip 安装脚本以来的 pywin32 `pip install pywin32`
3. 安装 win 最新版 wps 2019

#### 使用

代码里的例子 https://github.com/zjsxwc/wpsconverttopdf/blob/main/WpsConvertToPdf.py#L111
```python
x = WpsConvertToPdf()
x.convert(r"C:\Users\zjsxwc\Desktop\WpsConvertToPdf\test.xlsx")
```

