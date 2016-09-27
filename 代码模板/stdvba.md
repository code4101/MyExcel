## Enum 颜色表

```vb
Enum 颜色表
    标准字段颜色 = 15773696 'RGB(0, 176, 240)   蓝色
    额外字段颜色 = 5296274  'RGB(146, 208, 80)  绿色
    字段分类颜色 = 65535    'RGB(255, 255, 0)   黄色
End Enum
```

## QuickSort：快速排序

```vb
Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)

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
```
参考资料：smink, [VBA array sort function?](http://stackoverflow.com/questions/152319/vba-array-sort-function), stackoverflow, 2008.9