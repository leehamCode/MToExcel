# MToExcel

#### 介绍

一个把泛型T集合直接转化为Workbook的工具,主要用于简化CellStyle的样式代码.
<font color=red>提供Attribute设置样式,同时也提供自定义方法(Action)来提供定制NPOI代码.</font>
如果想要<b><u>简化NPOI样式代码或是需求对Excel样式的要求比较高,你可以尝试使用该工具.</u></b>

语言:C#
依赖:NPOI

#### 软件架构
/Converter
  ---BasicConverter
  ---WrapperConverter
  ---MTConverter
/Attribute
  ---ReferenceType
  ---IgnoreType
  ---HeaderName
  ---..........

#### 安装教程

现已上传NuGet
https://www.nuget.org/packages/MToExcel


#### 使用说明

文档网站现已上传github
https://leehamcode.github.io/mtoexcel_docs/

现在的使用很简单
1.  创建WrapperConverter对象
2.  调用ConverterToExcel<T>（List<T> list）方法，讲model集合传入
        可以更具ReferenceType标签指定要不要打印引用对象
        可以使用IgnoreType标签忽略掉不想要出现的属性
        可以使用HeadName标签指定属性名打印在Excel表头的列
        .............................................
3.  把返回的IWorkbook写到文件即可
代码的主方法中是测试用例，运行即可

#### 参与贡献

1.  罗马苏丹穆罕默德


#### 更新
随缘更新
