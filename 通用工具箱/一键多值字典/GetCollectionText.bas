Attribute VB_Name = "getCollectionText"
Function get����Text( _
    ByVal c As Collection, _
    Optional ���� As Boolean = True, _
    Optional �ظ������ As Boolean = True, _
    Optional ��ָ��� As String = ", ", _
    Optional �ظ�������ָ��� = "", _
    Optional ����1����ʾ As Boolean = False) As String
'(1)
    If �ظ������ Then
        Set c = ��Collection�ظ������(c, �ظ�������ָ���, ����1����ʾ)
    End If
'(2)
    If ���� Then
        Set c = ��Collection����(c)
    End If
'(3)
    Dim ans As String
    For Each cc In c
        If ans = "" Then
            ans = cc
        Else
            ans = ans & ��ָ��� & cc
        End If
    Next cc
'(4)return
    get����Text = ans
End Function
Private Function ��Collection����(c As Collection) As Collection
'ʵ��ԭ���Ƚ�����ת������Array����Arrayʹ��quicksort��Ȼ���ٽ��ź����Array����µ�Colllection
    Set ��Collection���� = New Collection
    Dim A() As Variant
    A = CollectionToArray(c)
    Call QuickSort(A, LBound(A), UBound(A))
    For Each aa In A
        ��Collection����.Add aa
    Next aa
End Function
Private Function ��Collection�ظ������(c As Collection, Optional �ظ�������ָ��� = "", Optional ����1����ʾ As Boolean = False) As Collection
'(1)�ȼ����ÿһ�������
    Dim cnt As New Dictionary '���ڽ����ֵ丨��
    For Each k In c
        cnt(k) = cnt(k) + 1
    Next k
'(2)����µļ���
    Dim d As New Dictionary
    Set ��Collection�ظ������ = New Collection
    For Each k In c
        If Not d.Exists(k) Then 'd�洢�Ѿ�visited����
            If ����1����ʾ And cnt(k) = 1 Then
                ��Collection�ظ������.Add k
            Else
                ��Collection�ظ������.Add k & �ظ�������ָ��� & cnt(k)
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

' ��Դ:http://stackoverflow.com/questions/152319/vba-array-sort-function
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
    c.Add "ƻ��"
    c.Add "ƻ��"
    c.Add "��"
    c.Add "�㽶"
    c.Add "��"
    
    '�����򣬼���
    Debug.Print get����Text(c, False)               'ƻ��2, ��2, �㽶1
    '���򣬼���
    Debug.Print get����Text(c)                      '��2, ƻ��2, �㽶1
    '������
    Debug.Print get����Text(c, �ظ������:=False)   '��, ��, ƻ��, ƻ��, �㽶
    '���ķָ����
    Debug.Print get����Text(c, ��ָ���:=";")       '��2;ƻ��2;�㽶1
End Sub
