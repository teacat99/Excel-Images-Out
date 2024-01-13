# Excel图片导出工具

## 项目描述

这是一个基于Python和ChatGPT开发的Excel图片导出工具。它专门用于从Excel文件中批量提取并重命名导出图片。此工具非常适合在使用Excel在线表格收集图片的场景中，提供了一种简单、高效的图片导出方法。

## 特点

- **批量导出**：支持从Excel文件中批量提取图片。
- **重命名功能**：导出的图片可以根据特定规则进行重命名，提高文件管理效率。
- **用户友好**：简洁的用户界面，易于操作。

## 系统要求

- 兼容Windows操作系统。
- 需要Python环境（可选，如果想要从源代码运行或修改工具）。

## 开始使用

1. 下载exe文件。
2. 双击运行程序。
3. 按照程序指示选择Excel文件，导出路径默认为Excel的文件位置。

## 测试案例

```python
欢迎使用Excel图片导出工具，请确保所有要导出的图片均为浮动格式，否则无法导出。

免责声明：图片导出前请注意Excel文件的备份，使用本程序造成的Excel文件损坏等，均由用户自行承担，与作者无关。

请提供Excel文件的完整路径，例如：C:\Users\YourName\Documents\example.xlsx，可以使用 Ctrl+Z 加回车退出输入。
请输入Excel文件的路径: D:\TeacatData\Desktop\打包\收货核验明细.xlsx
D:\TeacatData\Desktop\打包\收货核验明细.xlsx
基础输出文件夹设定为：D:\TeacatData\Desktop\打包

正在加载工作簿...

请选择图片所在sheet：
1. 导出所有sheet图片
2. sheet(汇总)
请输入选项（默认为1）：2


请选择图片命名方式：
1. 当前单元格位置
2. 当前单元格内容
3. 前一个单元格内容
请输入选项（默认为1）: 3


请选择图片导出格式：
1. PNG
2. JPG
请输入选项（默认为1）: 2


开始处理图像...

正在处理工作表：汇总

创建工作表文件夹：D:\TeacatData\Desktop\打包\汇总

发现图像位于：B2, 准备处理...
图像保存至：D:\TeacatData\Desktop\打包\汇总\20240113 设备现场照片.jpg

发现图像位于：B3, 准备处理...
图像保存至：D:\TeacatData\Desktop\打包\汇总\设备 0000000000000001 CPU.jpg

发现图像位于：B4, 准备处理...
图像保存至：D:\TeacatData\Desktop\打包\汇总\设备 0000000000000001 RAID卡.jpg

发现图像位于：B5, 准备处理...
图像保存至：D:\TeacatData\Desktop\打包\汇总\设备 0000000000000001 电源模块.jpg

发现图像位于：B6, 准备处理...
图像保存至：D:\TeacatData\Desktop\打包\汇总\设备 0000000000000001 风扇.jpg

发现图像位于：B7, 准备处理...
图像保存至：D:\TeacatData\Desktop\打包\汇总\设备 0000000000000001 内存.jpg

发现图像位于：B8, 准备处理...
图像保存至：D:\TeacatData\Desktop\打包\汇总\设备 0000000000000001 资产 W1234567890.jpg

发现图像位于：B9, 准备处理...
图像保存至：D:\TeacatData\Desktop\打包\汇总\设备 0000000000000001 Sku AB1.jpg

发现图像位于：B10, 准备处理...
图像保存至：D:\TeacatData\Desktop\打包\汇总\设备 0000000000000001 网卡.jpg

发现图像位于：B11, 准备处理...
图像保存至：D:\TeacatData\Desktop\打包\汇总\设备 0000000000000001 硬盘.jpg

发现图像位于：B12, 准备处理...
图像保存至：D:\TeacatData\Desktop\打包\汇总\设备 0000000000000001 主板.jpg

图像导出完成。输出目录：D:\TeacatData\Desktop\打包

按Enter键退出...
```

## 后记

没找到相关的工具，但是使用键盘宏真的费时费力，只能带着ChatGPT自己上了。所以..

该项目应运而生。它旨在解决使用Excel在线表格收集图片时的导出难题，帮助用户节省时间，提高效率。


## 联系方式

如果您有任何问题或建议，请通过以下方式联系我：

- GitHub: [teacat99](https://github.com/teacat99)
- 邮箱: teacat99@outlook.com

