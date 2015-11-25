0.背景说明
====
在ChinaTower工作期间，开发一些项目常用的函数、过程，独立写在一个模块，方便其他程序的开发。


2015/11/25
----
感觉大部分功能应该都整理到了，函数清单如下:
```vb
Function is广义抱杆(ByVal 铁塔类型 As String) As Boolean
Function 地址简称(ByVal 原名 As String) As String
Function 运营商(ByVal s As String) As String
Function 分类(ByVal 是否新增 As String, ByVal 是否存在 As String, ByVal 稽核结果 As String) As String
Function 塔型分类(ByVal 塔名 As String) As String
Function 场景分类(ByVal 所在地点 As String) As String
Function get站址编码(ByVal s As String, Optional ByVal times As Long = 1, Optional ByVal 索引超出返回空值 As Boolean = False) As String
Function get编码后的站名(ByVal s As String) As String
```
