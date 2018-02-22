## 简介

ExcelChart是基于xlsxwriter封装的库，让用户更加简单地创建Excel图表。



### 支持的图表类型

* 柱形图：column
* 条形图：bar
* 折线图：line
* 饼图：pie
* 雷达图：radar
* 散点图：scatter
* 面积图：area
* 圆环图：doughnut
* 股价图：stock





## 安装

### pip安装

该项目目前已传到pip上，在终端通过以下命令可以进行安装：

```
pip install excelchart
```

如果您没有安装pip，请查阅https://pip.pypa.io/en/stable/installing/获得详细的安装教程。



### 源码安装

下载项目文件，解压后找到setup.py文件目录，在终端切换到该目录后通过以下命令安装：

```
python setup.py install
```



## 创建Excel图表

为了方便处理和标准化，ExcelChart的输入数据是Pandas的DataFrame数据结构。另外，绘制后的图表会保存在新的Excel文件里面，原始数据也会写入在单元格中，图表可以选择插入单元格或者在图表Sheet中。



### 绘制柱形图

```python
import pandas as pd
from excelchart import ExcelChart

data = pd.read_excel('data.xlsx')
ec = ExcelChart('chart.xlsx')

chart = ec.bar(data)
ec.save()
```



### 绘制多个图表

```python
import pandas as pd
from excelchart import ExcelChart


```



## 设置图表参数

ExcelChart提供了几十个函数来设置图表参数，包括标题、图例、坐标轴、网格线等。



## 自定义模板

也许ExcelChart提供的默认图表



## 联系作者

如果您对该项目有任何意见或者提交bug，请发送邮件至jiada.yang@foxmail.com。