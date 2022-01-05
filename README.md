# MToExcel

#### 介绍
放一个自己写的把泛型集合写入Excel的小工具
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

目前只有一个基本的转化类和它的一个包装类
还有一个Attribute类用来标记引用类型


#### 安装教程

不需要安装

#### 使用说明
现在的使用很简单
1.  创建WrapperConverter对象
2.  调用ConverterToExcel<T>（List<T> list）方法，讲model集合传入
        可以更具ReferenceType标签指定要不要打印引用对象
        可以使用IgnoreType标签忽略掉不想要出现的属性
        可以使用HeadName标签指定属性名打印在Excel表头的列
3.  把返回的IWorkbook写到文件即可
代码的主方法中是测试用例，运行即可
#### 参与贡献

1.  罗马苏丹穆罕默德


#### 更新
随缘更新
