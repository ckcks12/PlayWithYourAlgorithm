Attribute VB_Name = "mod_Play"
Option Explicit

'----------------------------------Ű����
'=              ������ �������� ����
'+              ���׿� ������ ����
'-              ���׿� ������ ��
'*              ����
'/              ����
'^              ������ ���׸�ŭ �ŵ�����
'%              ������
'print A        A�μ� (��������)
'goto 2         2������ ������ �̵�
'if A b; c;     A�� ���ǹ��˻� -  ==���� !=�ٸ��� > >= < =< OR AND ������� A�� ���̸� b �����̸� c
'loop i=a b c;  i��������, a���� b���� 1�� ���� C�� �̵�.
'back a;b


'�������

Dim PB              As New PropertyBag
Dim Source()        As String
Dim Value           As New Collection
Dim POS             As Long

Sub Ready() 'ó������ÿ� PB���� ����Ÿ��������
    Dim LOFPoint       As Long
    Dim varTemp         As Variant
    Dim bArr()          As Byte
    Dim iFF             As Integer: iFF = FreeFile
    
    GoTo d
    Open App.Path & "\" & App.EXEName & ".exe" For Binary As #iFF
        Get #iFF, LOF(iFF) - 3, LOFPoint
        Seek #iFF, LOFPoint
        Get #iFF, , varTemp
        bArr = varTemp
        PB.Contents = bArr
    Close #iFF
    
    '---PB���� ���� ������ ����  �빮�ڴ� ȥ���� ����Ͽ� ������� �ʱ�
    
    'caption                ������
d:
    With PB
        Dim Temp() As String, i As Integer
        Temp = Split(Replace$(frm_Play.Text1.Text, vbCrLf, vbNullString), "||")
        For i = 0 To UBound(Temp())
            .WriteProperty i, Temp(i)
        Next i
        .WriteProperty "count", UBound(Temp)
    End With
    
    With PB
        frm_Play.Caption = .ReadProperty("caption", "Paly with your Algorithm")
        If Val(.ReadProperty("count")) > 0 And .ReadProperty("0") = "Begin" Then
            Call Begin
        Else
            MsgBox "�߸��� �ҽ�"
        End If
    End With
End Sub

Sub Begin() '�ҽ������غ��ϴ� ����
    ReDim Source(PB.ReadProperty("count")) As String
    Dim i               As Long
    Dim Temp()          As String
    
    Source(1) = PB.ReadProperty(1) '��������κ��� �׻� ù��°���ܿ�. ���⼭ �ʱ�ȭ��Ű��. �׻� ����θ� �ʱ�ȭ���Ѿ���.
    Temp = Split(Source(1), ";")
    For i = 0 To UBound(Temp) - 1&
        SaveValue Split(Temp(i), "=")(0), Split(Temp(i), "=")(1)
        DoEvents
    Next i
    
    If PB.ReadProperty(UBound(Source)) = "End" Then '�������� �ҽ�Ȯ���� ���� Begin�� End�� �˻���
        For i = 2 To UBound(Source) - 1
            Source(i) = PB.ReadProperty(i) '�ҽ��� ���ܺ��� ������
            DoEvents
        Next i
    Else
        MsgBox "�߸��� �ҽ�"
        Exit Sub
    End If
    
    POS = 2
    Do Until POS = Val(PB.ReadProperty("count"))
        ExeCute POS
    Loop
     
End Sub

Sub ExeCute(ByVal EXEPOS As Long)  '�ҽ������Ű�� ����
    'EXEPOS�� ���� �а��ִ� ��������ġ
    Dim This As String: This = Source(EXEPOS)
    Dim Temp As String
    Dim sLeft As String
    Dim sRight As String
    Dim LeftPos As Integer
    Dim RightPos As Integer
    Dim i       As Integer
    
    'ù��°�� ���깮���� ���ǹ������� �˻�
    If InStr(This, "if") Or InStr(This, "print") Or InStr(This, "loop") Or Left$(This, 1&) = ";" Then
        '���ǹ��̶��
        If Left$(This, 2&) = "if" Then
            Temp = Mid$(This, 4&, InStr(4&, This, " ") - 4&)
            If IsMark(Temp) Then
                '���� ���ǿ��� �������ʿ��Ұ��
                If InStr(Temp, "!=") Then
                    Temp = ControlUnit(Split(Temp, "!=")(0)) & "!=" & ControlUnit(Split(Temp, "!=")(1))
                ElseIf InStr(Temp, "<=") Then
                    Temp = ControlUnit(Split(Temp, "<=")(0)) & "<=" & ControlUnit(Split(Temp, "<=")(1))
                ElseIf InStr(Temp, ">=") Then
                    Temp = ControlUnit(Split(Temp, ">=")(0)) & ">=" & ControlUnit(Split(Temp, ">=")(1))
                ElseIf InStr(Temp, "=") Then
                    Temp = ControlUnit(Split(Temp, "=")(0)) & "=" & ControlUnit(Split(Temp, "=")(1))
                ElseIf InStr(Temp, "<") Then
                    Temp = ControlUnit(Split(Temp, "<")(0)) & "<" & ControlUnit(Split(Temp, "<")(1))
                ElseIf InStr(Temp, ">") Then
                    Temp = ControlUnit(Split(Temp, ">")(0)) & ">" & ControlUnit(Split(Temp, ">")(1))
                End If
                '����Ϸ�!
            End If
            If ALU(Temp) = True Then
                POS = Val(Mid$(This, InStr(This, "goto") + 5&, InStr(This, ";") - InStr(This, "goto") + 5&))
                'True ����
            Else
                POS = Val(Mid$(This, InStr(This, ";") + 7&, Len(This) - InStr(This, ";") + 7&))
                'False�κн���
            End If
        ElseIf Left$(This, 1&) = ";" Then 'goto��
            POS = Val(Mid$(This, 2))
        ElseIf Left$(This, 5&) = "print" Then 'print���̶�� print���� ;�ڿ� ���������ų�ڵ��ε��� ����.
            Temp = SplitEX(This, " ", ";")
            If IsNumber(Temp) Then
                MsgBox Val(Temp), vbOKOnly, "print"
            ElseIf IsMark(Temp) Then
                MsgBox ControlUnit(Temp), vbOKOnly, "print"
            Else
                MsgBox GetValue(Temp), vbOKOnly, "print"
            End If
            POS = Val(Mid$(This, InStr(This, ";") + 1&))
        ElseIf Left$(This, 4&) = "loop" Then 'LOOP��!
            Dim TempValue As String: TempValue = Mid$(This, 6&, InStr(This, "=") - 6&)
            Dim TempValue2 As String: TempValue2 = Split(Split(This, " ")(2), ";")(0)
            Call SaveValue(TempValue, SplitEX(This, "=", " "))
            Do
                If GetValue(TempValue) = TempValue2 Then Exit Do
                ExeCute (Val(Mid$(This, InStrRev(This, " ") + 1&)))
            Loop
        ElseIf Left$(This, 4&) = "back" Then 'BACK ��!!!!!
            This = Replace$(This, ";", vbNullString)
            Call SaveValue(Mid$(This, 6&), Val(GetValue(Mid$(This, 6&))) + 1&)
        End If
        '�ϴ���, ����Ʈ����� ���ǹ��������� ������ ������ �������վ����, ���߿��� ������ �ȵ�!
        '�� ����.. ControlUnit�� �߰��� �׳� �ٷ� �ؼ����� ������ ��¿�� ����Ʈ�� �׷����ۿ�..
    ElseIf InStr(This, "+") Or InStr(This, "-") Or InStr(This, "*") Or InStr(This, "/") Or InStr(This, "^") Or InStr(This, "%") Or InStr(This, "=") Then
        '�����̴� =�� �������װ�, ���׿��� �и��Ͽ�, ������ ALU���ְ� ���׺����� �����ϸ��. ������ �׻� ����
        '��������� ^ > * > / > % > + > -
        '������ ���帶������ �ϰ� ���׿��� ������������� ���׿����� �ǳ���....? �װ� ALU�������ִϱ� ..����
        
        sLeft = Replace$(Split(This, "=")(0&), " ", vbNullString)
        
        sRight = Left$(Replace$(Split(This, "=")(1&), " ", vbNullString), InStr(Replace$(Split(This, "=")(1&), " ", vbNullString), ";") - 1&)
     
        SaveValue sLeft, ControlUnit(sRight)
        POS = Val(Mid$(This, InStr(This, ";") + 1&))
    ElseIf This = "End" Then
        Exit Sub
    End If
End Sub

Function ALU(ByVal Expression As String) As String
    'Expression             ��
    '>, >=, =, <=, <, != �� ������ Logical Unit����
    '+, -, *, /, ^, % �� ������� Arithmetic Unit����
        
    Dim sLeft               As String '����
    Dim sRight              As String '����
    
    If InStr(Expression, ">=") Then
        sLeft = Split(Expression, ">=")(0)
        sRight = Split(Expression, ">=")(1)
        
        If Not IsNumber(sLeft) Then
            sLeft = GetValue(sLeft, "null")
            If sLeft = Null Then MsgBox "���� �ҷ����� ����" '���� �ҷ��������ʾ������
        End If
        If Not IsNumber(sRight) Then
            sRight = GetValue(sRight, "null")
            If sRight = Null Then MsgBox "���� �ҷ����� ����"
        End If
        
        If Val(sLeft) >= Val(sRight) Then
            ALU = True
        Else
            ALU = False
        End If
    ElseIf InStr(Expression, "<=") Then
        sLeft = Split(Expression, "<=")(0)
        sRight = Split(Expression, "<=")(1)
        
        If Not IsNumber(sLeft) Then
            sLeft = GetValue(sLeft, "null")
            If sLeft = Null Then MsgBox "���� �ҷ����� ����" '���� �ҷ��������ʾ������
        End If
        If Not IsNumber(sRight) Then
            sRight = GetValue(sRight, "null")
            If sRight = Null Then MsgBox "���� �ҷ����� ����"
        End If
        
        If Val(sLeft) <= Val(sRight) Then
            ALU = True
        Else
            ALU = False
        End If
    ElseIf InStr(Expression, ">") Then
        sLeft = Split(Expression, ">")(0)
        sRight = Split(Expression, ">")(1)
        
        If Not IsNumber(sLeft) Then
            sLeft = GetValue(sLeft, "null")
            If sLeft = Null Then MsgBox "���� �ҷ����� ����" '���� �ҷ��������ʾ������
        End If
        If Not IsNumber(sRight) Then
            sRight = GetValue(sRight, "null")
            If sRight = Null Then MsgBox "���� �ҷ����� ����"
        End If
        
        If Val(sLeft) > Val(sRight) Then
            ALU = True
        Else
            ALU = False
        End If
    ElseIf InStr(Expression, "!=") Then
        sLeft = Split(Expression, "!=")(0)
        sRight = Split(Expression, "!=")(1)
        
        If Not IsNumber(sLeft) Then
            sLeft = GetValue(sLeft, "null")
            If sLeft = Null Then MsgBox "���� �ҷ����� ����" '���� �ҷ��������ʾ������
        End If
        If Not IsNumber(sRight) Then
            sRight = GetValue(sRight, "null")
            If sRight = Null Then MsgBox "���� �ҷ����� ����"
        End If
        
        If Val(sLeft) <> Val(sRight) Then
            ALU = True
        Else
            ALU = False
        End If
    ElseIf InStr(Expression, "<") Then
        sLeft = Split(Expression, "<")(0)
        sRight = Split(Expression, "<")(1)
        
        If Not IsNumber(sLeft) Then
            sLeft = GetValue(sLeft, "null")
            If sLeft = Null Then MsgBox "���� �ҷ����� ����" '���� �ҷ��������ʾ������
        End If
        If Not IsNumber(sRight) Then
            sRight = GetValue(sRight, "null")
            If sRight = Null Then MsgBox "���� �ҷ����� ����"
        End If
        
        If Val(sLeft) < Val(sRight) Then
            ALU = True
        Else
            ALU = False
        End If
    ElseIf InStr(Expression, "=") Then
        sLeft = Split(Expression, "=")(0)
        sRight = Split(Expression, "=")(1)
        
        If Not IsNumber(sLeft) Then
            sLeft = GetValue(sLeft, "null")
            If sLeft = Null Then MsgBox "���� �ҷ����� ����" '���� �ҷ��������ʾ������
        End If
        If Not IsNumber(sRight) Then
            sRight = GetValue(sRight, "null")
            If sRight = Null Then MsgBox "���� �ҷ����� ����"
        End If
        
        If Val(sLeft) = Val(sRight) Then
            ALU = True
        Else
            ALU = False
        End If
    '-------------------------------------------------- ������� Logic Unit
    '-------------------------------------------------- ������� Arithmetic Unit
    ElseIf InStr(Expression, "+") Then
        sLeft = Split(Expression, "+")(0)
        sRight = Split(Expression, "+")(1)
        
        If Not IsNumber(sLeft) Then
            sLeft = GetValue(sLeft, "null")
            If sLeft = Null Then MsgBox "���� �ҷ����� ����" '���� �ҷ��������ʾ������
        End If
        If Not IsNumber(sRight) Then
            sRight = GetValue(sRight, "null")
            If sRight = Null Then MsgBox "���� �ҷ����� ����"
        End If
        
        ALU = Val(sLeft) + Val(sRight)
    ElseIf InStr(Expression, "-") Then
        sLeft = Split(Expression, "-")(0)
        sRight = Split(Expression, "-")(1)
        
        If Not IsNumber(sLeft) Then
            sLeft = GetValue(sLeft, "null")
            If sLeft = Null Then MsgBox "���� �ҷ����� ����" '���� �ҷ��������ʾ������
        End If
        If Not IsNumber(sRight) Then
            sRight = GetValue(sRight, "null")
            If sRight = Null Then MsgBox "���� �ҷ����� ����"
        End If
        
        ALU = Val(sLeft) - Val(sRight)
    ElseIf InStr(Expression, "*") Then
        sLeft = Split(Expression, "*")(0)
        sRight = Split(Expression, "*")(1)
        
        If Not IsNumber(sLeft) Then
            sLeft = GetValue(sLeft, "null")
            If sLeft = Null Then MsgBox "���� �ҷ����� ����" '���� �ҷ��������ʾ������
        End If
        If Not IsNumber(sRight) Then
            sRight = GetValue(sRight, "null")
            If sRight = Null Then MsgBox "���� �ҷ����� ����"
        End If
        
        ALU = Val(sLeft) * Val(sRight)
    ElseIf InStr(Expression, "/") Then
        sLeft = Split(Expression, "/")(0)
        sRight = Split(Expression, "/")(1)
        
        If Not IsNumber(sLeft) Then
            sLeft = GetValue(sLeft, "null")
            If sLeft = Null Then MsgBox "���� �ҷ����� ����" '���� �ҷ��������ʾ������
        End If
        If Not IsNumber(sRight) Then
            sRight = GetValue(sRight, "null")
            If sRight = Null Then MsgBox "���� �ҷ����� ����"
        End If
        
        ALU = Val(sLeft) / Val(sRight)
    ElseIf InStr(Expression, "^") Then
        sLeft = Split(Expression, "^")(0)
        sRight = Split(Expression, "^")(1)
        
        If Not IsNumber(sLeft) Then
            sLeft = GetValue(sLeft, "null")
            If sLeft = Null Then MsgBox "���� �ҷ����� ����" '���� �ҷ��������ʾ������
        End If
        If Not IsNumber(sRight) Then
            sRight = GetValue(sRight, "null")
            If sRight = Null Then MsgBox "���� �ҷ����� ����"
        End If
        
        ALU = Val(sLeft) ^ Val(sRight)
    ElseIf InStr(Expression, "%") Then
        sLeft = Split(Expression, "%")(0)
        sRight = Split(Expression, "%")(1)
        
        If Not IsNumber(sLeft) Then
            sLeft = GetValue(sLeft, "null")
            If sLeft = Null Then MsgBox "���� �ҷ����� ����" '���� �ҷ��������ʾ������
        End If
        If Not IsNumber(sRight) Then
            sRight = GetValue(sRight, "null")
            If sRight = Null Then MsgBox "���� �ҷ����� ����"
        End If
        
        ALU = Val(sLeft) Mod Val(sRight)
    End If
End Function

Function SplitEX(ByVal Str As String, ByVal sStart As String, ByVal sEnd As String) As String
    SplitEX = Split(Split(Str, sStart)(1), sEnd)(0)
End Function

Function IsNumber(ByVal Str As String) As Boolean
    Dim i               As Integer
    
    For i = 1 To Len(Str)
        If Asc(Mid$(Str, i, 1)) < 48 Or Asc(Mid$(Str, i, 1)) > 57 Then
            IsNumber = False
            Exit Function
        End If
    Next i
    
    IsNumber = True
End Function

Function SaveValue(ByVal Str As String, ByVal V As String)
    If IsValueExist(Str) Then Value.Remove (Str)
    Value.Add V, Str
End Function

Function GetValue(ByVal Str As String, Optional StrEX As String) As String
    If Not IsValueExist(Str) Then GetValue = StrEX: Exit Function
    GetValue = Value.Item(Str)
End Function

Function IsValueExist(ByVal Str As String) As Boolean
On Error GoTo E001
    Value.Item (Str)
    IsValueExist = True
    Exit Function
E001:
    IsValueExist = False
End Function

Function IsMark(ByVal Str As String) As Boolean
    If InStr(Str, "+") Or InStr(Str, "-") Or InStr(Str, "*") Or InStr(Str, "/") Or InStr(Str, "^") Or InStr(Str, "%") Then
        IsMark = True
    Else
        IsMark = False
    End If
End Function

Function ControlUnit(ByVal sRight As String) As String
    
    Dim i               As Long
    Dim LeftPos         As Long
    Dim RightPos        As Long
    Dim Temp            As String
    
    '����켱������ ���� ��ȣ��!^^
    
    Do Until InStr(sRight, ")") = 0
        LeftPos = InStrRev(sRight, "(")
        RightPos = InStr(sRight, ")")
        sRight = Left$(sRight, LeftPos - 1&) & ALU(Mid$(sRight, LeftPos + 1&, RightPos - LeftPos - 1&)) & Right$(sRight, Len(sRight) - RightPos)
        DoEvents
    Loop
    
    Do Until InStr(sRight, "^") = 0   ' ^�� ������������ ��� ġȯ�۾�
            LeftPos = InStr(sRight, "^")
            For i = LeftPos - 1& To 1& Step -1
                Temp = Mid$(sRight, i, 1&)
                If Temp = "^" Or Temp = "/" Or Temp = "*" Or Temp = "%" Or Temp = "+" Or Temp = "-" Then
                    LeftPos = i + 1&
                    Exit For
                End If
                DoEvents
            Next i
            If i = 0& Then LeftPos = 1&
            
            RightPos = InStr(sRight, "^")
            For i = RightPos + 1& To Len(sRight) - 1&
                Temp = Mid$(sRight, i, 1&)
                If Temp = "^" Or Temp = "/" Or Temp = "*" Or Temp = "%" Or Temp = "+" Or Temp = "-" Then
                    RightPos = i - 1&
                    Exit For
                End If
                DoEvents
            Next i
            If i = Len(sRight) Then RightPos = i
     
            sRight = Left$(sRight, LeftPos - 1&) & ALU(Mid$(sRight, LeftPos, RightPos - LeftPos + 1&)) & Right$(sRight, Len(sRight) - RightPos)
        Loop
        
        Do Until InStr(sRight, "*") = 0   '*�� ������������ ��� ġȯ�۾�
            LeftPos = InStr(sRight, "*")
            For i = LeftPos - 1& To 1& Step -1
                Temp = Mid$(sRight, i, 1&)
                If Temp = "^" Or Temp = "/" Or Temp = "*" Or Temp = "%" Or Temp = "+" Or Temp = "-" Then
                    LeftPos = i + 1&
                    Exit For
                End If
                DoEvents
            Next i
            If i = 0& Then LeftPos = 1&
            
            RightPos = InStr(sRight, "*")
            For i = RightPos + 1& To Len(sRight) - 1&
                Temp = Mid$(sRight, i, 1&)
                If Temp = "^" Or Temp = "/" Or Temp = "*" Or Temp = "%" Or Temp = "+" Or Temp = "-" Then
                    RightPos = i - 1&
                    Exit For
                End If
                DoEvents
            Next i
            If i = Len(sRight) Then RightPos = i
     
            sRight = Left$(sRight, LeftPos - 1&) & ALU(Mid$(sRight, LeftPos, RightPos - LeftPos + 1&)) & Right$(sRight, Len(sRight) - RightPos)
        Loop
        
        Do Until InStr(sRight, "/") = 0   ' /�� ������������ ��� ġȯ�۾�
            LeftPos = InStr(sRight, "/")
            For i = LeftPos - 1& To 1& Step -1
                Temp = Mid$(sRight, i, 1&)
                If Temp = "^" Or Temp = "/" Or Temp = "*" Or Temp = "%" Or Temp = "+" Or Temp = "-" Then
                    LeftPos = i + 1&
                    Exit For
                End If
                DoEvents
            Next i
            If i = 0& Then LeftPos = 1&
            
            RightPos = InStr(sRight, "/")
            For i = RightPos + 1& To Len(sRight) - 1&
                Temp = Mid$(sRight, i, 1&)
                If Temp = "^" Or Temp = "/" Or Temp = "*" Or Temp = "%" Or Temp = "+" Or Temp = "-" Then
                    RightPos = i - 1&
                    Exit For
                End If
                DoEvents
            Next i
            If i = Len(sRight) Then RightPos = i
     
            sRight = Left$(sRight, LeftPos - 1&) & ALU(Mid$(sRight, LeftPos, RightPos - LeftPos + 1&)) & Right$(sRight, Len(sRight) - RightPos)
        Loop
        
        Do Until InStr(sRight, "%") = 0   ' %�� ������������ ��� ġȯ�۾�
            LeftPos = InStr(sRight, "%")
            For i = LeftPos - 1& To 1& Step -1
                Temp = Mid$(sRight, i, 1&)
                If Temp = "^" Or Temp = "/" Or Temp = "*" Or Temp = "%" Or Temp = "+" Or Temp = "-" Then
                    LeftPos = i + 1&
                    Exit For
                End If
                DoEvents
            Next i
            If i = 0& Then LeftPos = 1&
            
            RightPos = InStr(sRight, "%")
            For i = RightPos + 1& To Len(sRight) - 1&
                Temp = Mid$(sRight, i, 1&)
                If Temp = "^" Or Temp = "/" Or Temp = "*" Or Temp = "%" Or Temp = "+" Or Temp = "-" Then
                    RightPos = i - 1&
                    Exit For
                End If
                DoEvents
            Next i
            If i = Len(sRight) Then RightPos = i
     
            sRight = Left$(sRight, LeftPos - 1&) & ALU(Mid$(sRight, LeftPos, RightPos - LeftPos + 1&)) & Right$(sRight, Len(sRight) - RightPos)
        Loop
        
        Do Until InStr(sRight, "+") = 0   ' +�� ������������ ��� ġȯ�۾�
            LeftPos = InStr(sRight, "+")
            For i = LeftPos - 1& To 1& Step -1
                Temp = Mid$(sRight, i, 1&)
                If Temp = "^" Or Temp = "/" Or Temp = "*" Or Temp = "%" Or Temp = "+" Or Temp = "-" Then
                    LeftPos = i + 1&
                    Exit For
                End If
                DoEvents
            Next i
            If i = 0& Then LeftPos = 1&
            
            RightPos = InStr(sRight, "+")
            For i = RightPos + 1& To Len(sRight) - 1&
                Temp = Mid$(sRight, i, 1&)
                If Temp = "^" Or Temp = "/" Or Temp = "*" Or Temp = "%" Or Temp = "+" Or Temp = "-" Then
                    RightPos = i - 1&
                    Exit For
                End If
                DoEvents
            Next i
            If i = Len(sRight) Then RightPos = i
     
            sRight = Left$(sRight, LeftPos - 1&) & ALU(Mid$(sRight, LeftPos, RightPos - LeftPos + 1&)) & Right$(sRight, Len(sRight) - RightPos)
        Loop
        
        Do Until InStr(sRight, "-") = 0   ' -�� ������������ ��� ġȯ�۾�
            '�����̿����׷��ɼ���������..
            If Val(sRight) < 0 Then Exit Do
            LeftPos = InStr(sRight, "-")
            For i = LeftPos - 1& To 1& Step -1
                Temp = Mid$(sRight, i, 1&)
                If Temp = "^" Or Temp = "/" Or Temp = "*" Or Temp = "%" Or Temp = "+" Or Temp = "-" Then
                    LeftPos = i + 1&
                    Exit For
                End If
                DoEvents
            Next i
            If i = 0& Then LeftPos = 1&
            
            RightPos = InStr(sRight, "-")
            For i = RightPos + 1& To Len(sRight) - 1&
                Temp = Mid$(sRight, i, 1&)
                If Temp = "^" Or Temp = "/" Or Temp = "*" Or Temp = "%" Or Temp = "+" Or Temp = "-" Then
                    RightPos = i - 1&
                    Exit For
                End If
                DoEvents
            Next i
            If i = Len(sRight) Then RightPos = i
    
            sRight = Left$(sRight, LeftPos - 1&) & ALU(Mid$(sRight, LeftPos, RightPos - LeftPos + 1&)) & Right$(sRight, Len(sRight) - RightPos)
        Loop
        
        ControlUnit = sRight
End Function
