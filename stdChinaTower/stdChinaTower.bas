Attribute VB_Name = "stdChinaTower"
'���������п���������Щ���
'Microsoft Scripting Runtime
'Microsoft VBScript Regular Expressions 5.5


Function is���屧��(ByVal �������� As String) As Boolean
    ' ö�ٱ����࣬û���ֵ�����Ϊ������������
    If InStr(��������, "����") Then
        ' �����������ֱ��˵ģ����ˡ����ˣ�4G�������汧�ˡ����ϱ��ˡ����߱��ˡ��ݶ�����
        is���屧�� = True
    ElseIf InStr(��������, "����") Then
        ' �磺�������ߡ���������
        is���屧�� = True
    ElseIf InStr(��������, "����") Then
        ' �磺��������
        is���屧�� = True
    ElseIf InStr(��������, "H��") Then
        is���屧�� = True
    ElseIf InStr(��������, "��ǽ��") Then
        is���屧�� = True
    ElseIf InStr(��������, "����") Then
        is���屧�� = True
    ElseIf InStr(��������, "�ڹ�") Then '�ʼ���2015��04��28�վͼ��ˣ����Ǵ���5��20�ղŲ���~~
        is���屧�� = True
    Else
        is���屧�� = False
    End If
End Function


'��ʱ��Ҫ���е�ַģ��ƥ�䣬��vlookupֻ���ڶ�key���о�ȷ��ƥ��
'��ʱɾ��һЩ�������ۣ��ܴ����߿�ƥ���������ƽ��ƥ����ȷ�����Լ���ΰ�����
Function ��ַ���(ByVal ԭ�� As String) As String
    '(1) ɾ����������
    ��ַ��� = Replace(ԭ��, "����", "")
    
    '(2)ɾ����ַ��6����������
    ��ַ��� = Replace(��ַ���, "˼��", "")
    ��ַ��� = Replace(��ַ���, "����", "")
    ��ַ��� = Replace(��ַ���, "����", "")
    ��ַ��� = Replace(��ַ���, "����", "")
    ��ַ��� = Replace(��ַ���, "ͬ��", "")
    ��ַ��� = Replace(��ַ���, "�谲", "")
    
    '(3)��һ���滻
    ��ַ��� = Replace(��ַ���, "��", "")
    ��ַ��� = Replace(��ַ���, "վ��", "")
    ��ַ��� = Replace(��ַ���, "��", "")
End Function

'��ʱ��������Ӫ��һ�����ʱ��Ҫ��������վַ��Ż�����Ӫ��
Function ��Ӫ��(ByVal s As String) As String
    If s Like "*����*" Then
        ��Ӫ�� = "����"
    ElseIf s Like "*��ͨ*" Then
        ��Ӫ�� = "��ͨ"
    ElseIf s Like "*�ƶ�*" Then
        ��Ӫ�� = "�ƶ�"
    Else
        ��Ӫ�� = "����"
    End If
End Function


'�ʲ������ڼ�ͨ�õĺ���
Function ����(ByVal �Ƿ����� As String, ByVal �Ƿ���� As String, ByVal ���˽�� As String) As String
    If ���˽�� = "ͨ��" Then
        If �Ƿ���� = "��" Then
            ���� = "�̿�"
        ElseIf �Ƿ����� = "��" Then
            ���� = "����"
        Else    '��������
            ���� = "���"
        End If
    ElseIf ���˽�� = "���©��д" Then
        ���� = "����"
    Else
        ���� = "��ȷ��"
    End If
End Function

Function ���ͷ���(ByVal ���� As String) As String
'   (���� ¥����)
'������
'������
'�Ĺ���
'H��
'�������߼�
'�Ǹ���
'������
'·�Ƹ�
'������
'��Яʽ����һ�廯
'����������
'   (����  ������)
'��������Φ��
'�������߼� (��¥��б��)
'����
'���澰���� (��������?����һ�����ߵ�)
'����
    If InStr(����, "����") Then
        ���ͷ��� = "������"
    ElseIf InStr(����, "����") Then
        ���ͷ��� = "¥����"
    ElseIf InStr(����, "����") Then
        ���ͷ��� = "������"
    ElseIf InStr(����, "����") Then
        ���ͷ��� = "������"
    Else
        ���ͷ��� = "¥����"
    End If
End Function

Function ��������(ByVal ���ڵص� As String) As String
'����û��ũ�壬ֻ��ֳ� �����С��͡�ũ�塱
    If InStr(���ڵص�, "˼��") Or InStr(���ڵص�, "����") Then
        �������� = "����"
    ElseIf InStr(���ڵص�, "ͬ��") Or InStr(���ڵص�, "�谲") Then
        �������� = "����"
    ElseIf InStr(���ڵص�, "�ֵ�") Then
        �������� = "����"
    ElseIf InStr(���ڵص�, "��") Then   'ǰ�Ĳ�Ӧ���ܽ��80%��������
        �������� = "����"
    '���������Ȼ�ܷ�����������Ϊ���Է���һд�˵�����ȴû�зֵֽ�����
    ElseIf InStr(���ڵص�, "���") Or InStr(���ڵص�, "��Ϫ") Or InStr(���ڵص�, "����") Then
        �������� = "����"
    ElseIf InStr(���ڵص�, "��Ӣ") Or InStr(���ڵص�, "����") Or InStr(���ڵص�, "�ӱ�") Then   '�����ļ����ֵ�����ʵ���и������ֵ�
        �������� = "����"
    ElseIf InStr(���ڵص�, "����") Or InStr(���ڵص�, "����") Then
        �������� = "����"
    Else
        �������� = "��ȷ��"
    End If
End Function


'�������ݿ�����t0924
's��Ҫ���ҵ�����
'times�Ƿ��صڼ���ƥ��Ľ�����������ƥ�������򷵻����һ��ƥ����
'����������ˡ������������ؿ�ֵ�������ٷ������һ��ƥ����
'��������ƥ��ɹ��ı��ʽ
'��ƥ�������ؿմ�
Function getվַ����(ByVal s As String, Optional ByVal times As Long = 1, Optional ByVal �����������ؿ�ֵ As Boolean = False) As String
'(1)��������:ʹ��staticֻ����һ��
    Static regx As RegExp
    If regx Is Nothing Then
        Set regx = New RegExp
        With regx
            .Pattern = "(?:����){0,1}..��������(?:����|����)(?:����){0,1}(?:\d{6})"
            .Global = True
        End With
    End If

'(2)ƥ�����
    Set mh = regx.Execute(s)
    
'(3)����ֵ
    'Debug.Print mh.Count
    '(a)���û��ƥ���������ؿ�
    If mh.Count = 0 Then
        getվַ���� = ""
        Exit Function
    End If
    '(b)ƥ�����ǿ�
    If mh.Count >= times Then
        getվַ���� = mh(times - 1).Value   '�����Ǵ�0��ʼ��ŵġ�����
    ElseIf �����������ؿ�ֵ Then
        getվַ���� = ""
    Else
        getվַ���� = mh(mh.Count - 1).Value
    End If
End Function

'��һ���ַ����У�ƥ��6�����ֺ�(�Դ�Ϊվַ����)����ȡ����" "��"+"���������
Function get������վ��(ByVal s As String) As String
'(1)��������:ʹ��staticֻ����һ��
    Static regx As RegExp
    If regx Is Nothing Then
        Set regx = New RegExp
        regx.Pattern = "(?:\d{6}\+?)(.*)"
    End If
'(2)ƥ�����
    Set mh = regx.Execute(s)
'(3)����ֵ
    If mh.Count > 0 Then get������վ�� = mh(0).SubMatches(0)
End Function

