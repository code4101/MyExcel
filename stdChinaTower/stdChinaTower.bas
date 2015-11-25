Attribute VB_Name = "stdChinaTower"
'请在引用中开启下面这些组件
'Microsoft Scripting Runtime
'Microsoft VBScript Regular Expressions 5.5


Function is广义抱杆(ByVal 铁塔类型 As String) As Boolean
    ' 枚举抱杆类，没出现的则认为是狭义铁塔类
    If InStr(铁塔类型, "抱杆") Then
        ' 类型名里面又抱杆的：抱杆、抱杆（4G）、地面抱杆、塔上抱杆、天线抱杆、屋顶抱杆
        is广义抱杆 = True
    ElseIf InStr(铁塔类型, "美化") Then
        ' 如：美化天线、美化外罩
        is广义抱杆 = True
    ElseIf InStr(铁塔类型, "天线") Then
        ' 如：美化天线
        is广义抱杆 = True
    ElseIf InStr(铁塔类型, "H杆") Then
        is广义抱杆 = True
    ElseIf InStr(铁塔类型, "附墙杆") Then
        is广义抱杆 = True
    ElseIf InStr(铁塔类型, "立杆") Then
        is广义抱杆 = True
    ElseIf InStr(铁塔类型, "壁挂") Then '笔记里2015年04月28日就加了，但是代码5月20日才补上~~
        is广义抱杆 = True
    Else
        is广义抱杆 = False
    End If
End Function


'有时候要进行地址模糊匹配，而vlookup只能在对key进行精确的匹配
'此时删掉一些干扰字眼，能大大提高可匹配量，如果平衡匹配正确率则看自己如何把握了
Function 地址简称(ByVal 原名 As String) As String
    '(1) 删掉厦门字眼
    地址简称 = Replace(原名, "厦门", "")
    
    '(2)删掉地址里6个区的名称
    地址简称 = Replace(地址简称, "思明", "")
    地址简称 = Replace(地址简称, "湖里", "")
    地址简称 = Replace(地址简称, "集美", "")
    地址简称 = Replace(地址简称, "海沧", "")
    地址简称 = Replace(地址简称, "同安", "")
    地址简称 = Replace(地址简称, "翔安", "")
    
    '(3)进一步替换
    地址简称 = Replace(地址简称, "区", "")
    地址简称 = Replace(地址简称, "站点", "")
    地址简称 = Replace(地址简称, "局", "")
End Function

'有时候三家运营商一起分析时，要根据物理站址编号划分运营商
Function 运营商(ByVal s As String) As String
    If s Like "*电信*" Then
        运营商 = "电信"
    ElseIf s Like "*联通*" Then
        运营商 = "联通"
    ElseIf s Like "*移动*" Then
        运营商 = "移动"
    Else
        运营商 = "错误"
    End If
End Function


'资产稽核期间通用的函数
Function 分类(ByVal 是否新增 As String, ByVal 是否存在 As String, ByVal 稽核结果 As String) As String
    If 稽核结果 = "通过" Then
        If 是否存在 = "否" Then
            分类 = "盘亏"
        ElseIf 是否新增 = "是" Then
            分类 = "盘增"
        Else    '不是新增
            分类 = "相符"
        End If
    ElseIf 稽核结果 = "表格漏填写" Then
        分类 = "盘增"
    Else
        分类 = "待确认"
    End If
End Function

Function 塔型分类(ByVal 塔名 As String) As String
'   (以下 楼面塔)
'单管塔
'三管塔
'四管塔
'H杆
'地面增高架
'角钢塔
'景观塔
'路灯杆
'仿生树
'便携式塔房一体化
'地面拉线塔
'   (以下  屋面塔)
'屋面拉线桅杆
'屋面增高架 (含楼顶斜撑)
'抱杆
'屋面景观塔 (含美化罩?美化一体天线等)
'其他
    If InStr(塔名, "屋面") Then
        塔型分类 = "屋面塔"
    ElseIf InStr(塔名, "地面") Then
        塔型分类 = "楼面塔"
    ElseIf InStr(塔名, "抱杆") Then
        塔型分类 = "屋面塔"
    ElseIf InStr(塔名, "其他") Then
        塔型分类 = "屋面塔"
    Else
        塔型分类 = "楼面塔"
    End If
End Function

Function 场景分类(ByVal 所在地点 As String) As String
'厦门没有农村，只需分出 “城市”和“农村”
    If InStr(所在地点, "思明") Or InStr(所在地点, "湖里") Then
        场景分类 = "城市"
    ElseIf InStr(所在地点, "同安") Or InStr(所在地点, "翔安") Then
        场景分类 = "乡镇"
    ElseIf InStr(所在地点, "街道") Then
        场景分类 = "城市"
    ElseIf InStr(所在地点, "镇") Then   '前四步应该能解决80%的问题了
        场景分类 = "乡镇"
    '后面代码虽然很繁琐，不过是为了以防万一写了地名，却没有分街道、镇
    ElseIf InStr(所在地点, "灌口") Or InStr(所在地点, "后溪") Or InStr(所在地点, "东孚") Then
        场景分类 = "乡镇"
    ElseIf InStr(所在地点, "侨英") Or InStr(所在地点, "杏林") Or InStr(所在地点, "杏滨") Then   '集美的几个街道，其实还有个集美街道
        场景分类 = "城市"
    ElseIf InStr(所在地点, "海沧") Or InStr(所在地点, "新阳") Then
        场景分类 = "城市"
    Else
        场景分类 = "不确定"
    End If
End Function


'基本内容开发于t0924
's是要查找的内容
'times是返回第几个匹配的结果，如果超过匹配数，则返回最后一个匹配结果
'但如果设置了“索引超出返回空值”，则不再返回最后一个匹配结果
'函数返回匹配成功的表达式
'无匹配结果返回空串
Function get站址编码(ByVal s As String, Optional ByVal times As Long = 1, Optional ByVal 索引超出返回空值 As Boolean = False) As String
'(1)正则设置:使用static只编译一次
    Static regx As RegExp
    If regx Is Nothing Then
        Set regx = New RegExp
        With regx
            .Pattern = "(?:二次){0,1}..福建厦门(?:自有|租赁)(?:二次){0,1}(?:\d{6})"
            .Global = True
        End With
    End If

'(2)匹配查找
    Set mh = regx.Execute(s)
    
'(3)返回值
    'Debug.Print mh.Count
    '(a)如果没有匹配结果，返回空
    If mh.Count = 0 Then
        get站址编码 = ""
        Exit Function
    End If
    '(b)匹配结果非空
    If mh.Count >= times Then
        get站址编码 = mh(times - 1).Value   '擦，是从0开始编号的。。。
    ElseIf 索引超出返回空值 Then
        get站址编码 = ""
    Else
        get站址编码 = mh(mh.Count - 1).Value
    End If
End Function

'在一个字符串中，匹配6个数字后(以此为站址编码)，读取增加" "或"+"后面的内容
Function get编码后的站名(ByVal s As String) As String
'(1)正则设置:使用static只编译一次
    Static regx As RegExp
    If regx Is Nothing Then
        Set regx = New RegExp
        regx.Pattern = "(?:\d{6}\+?)(.*)"
    End If
'(2)匹配查找
    Set mh = regx.Execute(s)
'(3)返回值
    If mh.Count > 0 Then get编码后的站名 = mh(0).SubMatches(0)
End Function

