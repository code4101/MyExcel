Attribute VB_Name = "getCollectionText"
Function get集合Text( _
    ByVal c As Collection, _
    Optional 排序 As Boolean = True, _
    Optional 重复项计数 As Boolean = True, _
    Optional 项分隔符 As String = ", ", _
    Optional 重复项计数分隔符 = "", _
    Optional 数量1不显示 As Boolean = False) As String
'(1)
    If 重复项计数 Then
        Set c = 对Collection重复项汇总(c, 重复项计数分隔符, 数量1不显示)
    End If
'(2)
    If 排序 Then
        Set c = 对Collection排序(c)
    End If
'(3)
    Dim ans As String
    For Each cc In c
        If ans = "" Then
            ans = cc
        Else
            ans = ans & 项分隔符 & cc
        End If
    Next cc
'(4)return
    get集合Text = ans
End Function
Private Function 对Collection排序(c As Collection) As Collection
'实现原理：先将集合转成数组Array，对Array使用quicksort，然后再将排好序的Array存回新的Colllection
    Set 对Collection排序 = New Collection
    Dim A() As Variant
    A = CollectionToArray(c)
    Call QuickSort(A, LBound(A), UBound(A))
    For Each aa In A
        对Collection排序.Add aa
    Next aa
End Function
Private Function 对Collection重复项汇总(c As Collection, Optional 重复项计数分隔符 = "", Optional 数量1不显示 As Boolean = False) As Collection
'(1)先计算出每一项的数量
    Dim cnt As New Dictionary '用于建立字典辅助
    For Each k In c
        cnt(k) = cnt(k) + 1
    Next k
'(2)算出新的集合
    Dim d As New Dictionary
    Set 对Collection重复项汇总 = New Collection
    For Each k In c
        If Not d.Exists(k) Then 'd存储已经visited的项
            If 数量1不显示 And cnt(k) = 1 Then
                对Collection重复项汇总.Add k
            Else
                对Collection重复项汇总.Add k & 重复项计数分隔符 & cnt(k)
            End If
            d.Add k, ""
        End If
    Next k
End Function
'https://brettdotnet.wordpress.com/2012/03/30/convert-a-collection-to-an-array-vba/
Private Function CollectionToArray(c As Collection) As Variant()
    Dim A() As Variant: ReDim A(1 To c.Count)
    Dim i As Long
    For i = 1 To c.Count
        A(i) = c.Item(i)
    Next
    CollectionToArray = A
End Function

' 来源:http://stackoverflow.com/questions/152319/vba-array-sort-function
Private Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)

  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)

     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If

  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi

End Sub


Private Sub Test()
    Dim c As New Collection
    c.Add "苹果"
    c.Add "苹果"
    c.Add "梨"
    c.Add "香蕉"
    c.Add "梨"
    
    '不排序，计数
    Debug.Print get集合Text(c, False)               '苹果2, 梨2, 香蕉1
    '排序，计数
    Debug.Print get集合Text(c)                      '梨2, 苹果2, 香蕉1
    '不计数
    Debug.Print get集合Text(c, 重复项计数:=False)   '梨, 梨, 苹果, 苹果, 香蕉
    '更改分割符号
    Debug.Print get集合Text(c, 项分隔符:=";")       '梨2;苹果2;香蕉1
End Sub
