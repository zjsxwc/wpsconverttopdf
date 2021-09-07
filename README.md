一个python脚本，用于自动把各种wps支持的文档，借助windows下的wps自动把各种文档转换为pdf文档，

#### 安装
1. 安装最新版的 python3 ，安装时记得把 python 放到系统 Path 里面勾选
2. 使用 pip 安装脚本依赖的 pywin32 `pip install pywin32`
3. 安装 win 最新版 wps 2019
4. 设置 wps 为 xlsx、docx 等拓展名文件的默认打开程序
5. 不建议用 win10 系统，我是虚拟机 win7

#### 使用

代码里的例子 https://github.com/zjsxwc/wpsconverttopdf/blob/main/WpsConvertToPdf.py#L111
```python
x = WpsConvertToPdf()
x.convert(r"C:\Users\zjsxwc\Desktop\WpsConvertToPdf\test.xlsx")
```

#### todo

由于wps对pptx的交互需要鼠标介入，
所以目前脚本还不支持ppt转pdf，
我也懒得写了，就是对ppt用鼠标特殊处理，
pywin32怎么模拟鼠标操作代码网上都是，很简单，
有需求的自己调吧。
