Attribute VB_Name = "mod_Play"
Option Explicit

'----------------------------------키워드
'=              우항을 좌항으로 대입
'+              우항에 좌항을 더함
'-              우항에 좌항을 뺌
'*              곱함
'/              나눔
'^              우항을 좌항만큼 거듭제곱
'%              나머지
'print A        A인쇄 (결과값출력)
'goto 2         2번으로 포인터 이동
'if A b; c;     A식 조건문검사 -  ==같다 !=다르다 > >= < =< OR AND 연산까지 A가 참이면 b 거짓이면 c
'loop i=a b c;  i변수선언, a부터 b까지 1씩 더함 C로 이동.
'back a;b


'실행모듈들

Dim PB              As New PropertyBag
Dim Source()        As String
Dim Value           As New Collection
Dim POS             As Long

Sub Ready() '처음실행시에 PB에서 데이타가져오기
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
    
    '---PB에서 직접 꺼내다 쓰기  대문자는 혼동을 우려하여 사용하지 않기
    
    'caption                폼제목
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
            MsgBox "잘못된 소스"
        End If
    End With
End Sub

Sub Begin() '소스실행준비하는 서브
    ReDim Source(PB.ReadProperty("count")) As String
    Dim i               As Long
    Dim Temp()          As String
    
    Source(1) = PB.ReadProperty(1) '변수선언부분은 항상 첫번째스텝에. 여기서 초기화시키기. 항상 상수로만 초기화시켜야함.
    Temp = Split(Source(1), ";")
    For i = 0 To UBound(Temp) - 1&
        SaveValue Split(Temp(i), "=")(0), Split(Temp(i), "=")(1)
        DoEvents
    Next i
    
    If PB.ReadProperty(UBound(Source)) = "End" Then '정상적인 소스확인을 위해 Begin과 End를 검사함
        For i = 2 To UBound(Source) - 1
            Source(i) = PB.ReadProperty(i) '소스를 스텝별로 가져옴
            DoEvents
        Next i
    Else
        MsgBox "잘못된 소스"
        Exit Sub
    End If
    
    POS = 2
    Do Until POS = Val(PB.ReadProperty("count"))
        ExeCute POS
    Loop
     
End Sub

Sub ExeCute(ByVal EXEPOS As Long)  '소스실행시키는 서브
    'EXEPOS는 현재 읽고있는 포인터위치
    Dim This As String: This = Source(EXEPOS)
    Dim Temp As String
    Dim sLeft As String
    Dim sRight As String
    Dim LeftPos As Integer
    Dim RightPos As Integer
    Dim i       As Integer
    
    '첫번째로 연산문인지 조건문인지를 검사
    If InStr(This, "if") Or InStr(This, "print") Or InStr(This, "loop") Or Left$(This, 1&) = ";" Then
        '조건문이라면
        If Left$(This, 2&) = "if" Then
            Temp = Mid$(This, 4&, InStr(4&, This, " ") - 4&)
            If IsMark(Temp) Then
                '만약 조건에서 연산이필요할경우
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
                '연산완료!
            End If
            If ALU(Temp) = True Then
                POS = Val(Mid$(This, InStr(This, "goto") + 5&, InStr(This, ";") - InStr(This, "goto") + 5&))
                'True 실행
            Else
                POS = Val(Mid$(This, InStr(This, ";") + 7&, Len(This) - InStr(This, ";") + 7&))
                'False부분실행
            End If
        ElseIf Left$(This, 1&) = ";" Then 'goto문
            POS = Val(Mid$(This, 2))
        ElseIf Left$(This, 5&) = "print" Then 'print문이라면 print문도 ;뒤에 다음실행시킬코드인덱스 붙임.
            Temp = SplitEX(This, " ", ";")
            If IsNumber(Temp) Then
                MsgBox Val(Temp), vbOKOnly, "print"
            ElseIf IsMark(Temp) Then
                MsgBox ControlUnit(Temp), vbOKOnly, "print"
            Else
                MsgBox GetValue(Temp), vbOKOnly, "print"
            End If
            POS = Val(Mid$(This, InStr(This, ";") + 1&))
        ElseIf Left$(This, 4&) = "loop" Then 'LOOP문!
            Dim TempValue As String: TempValue = Mid$(This, 6&, InStr(This, "=") - 6&)
            Dim TempValue2 As String: TempValue2 = Split(Split(This, " ")(2), ";")(0)
            Call SaveValue(TempValue, SplitEX(This, "=", " "))
            Do
                If GetValue(TempValue) = TempValue2 Then Exit Do
                ExeCute (Val(Mid$(This, InStrRev(This, " ") + 1&)))
            Loop
        ElseIf Left$(This, 4&) = "back" Then 'BACK 문!!!!!
            This = Replace$(This, ";", vbNullString)
            Call SaveValue(Mid$(This, 6&), Val(GetValue(Mid$(This, 6&))) + 1&)
        End If
        '일단은, 프린트라던지 조건문에서들은 무조건 변수는 변수만잇어야함, 식중에서 연산은 안됨!
        '은 무슨.. ControlUnit의 추가로 그냥 바로 해석가능 ㅋㅋㅋ 어쩔껴 프린트도 그럴수밖에..
    ElseIf InStr(This, "+") Or InStr(This, "-") Or InStr(This, "*") Or InStr(This, "/") Or InStr(This, "^") Or InStr(This, "%") Or InStr(This, "=") Then
        '연산이니 =는 꼭있을테고, 좌항우항 분리하여, 우항을 ALU에넣고 좌항변수에 대입하면됨. 좌항은 항상 변수
        '연산순서는 ^ > * > / > % > + > -
        '대입은 가장마지막에 하고 좌항에서 연산순서에따라서 좌항우항을 또나눌....? 그건 ALU에서해주니깐 ..ㅋㅋ
        
        sLeft = Replace$(Split(This, "=")(0&), " ", vbNullString)
        
        sRight = Left$(Replace$(Split(This, "=")(1&), " ", vbNullString), InStr(Replace$(Split(This, "=")(1&), " ", vbNullString), ";") - 1&)
     
        SaveValue sLeft, ControlUnit(sRight)
        POS = Val(Mid$(This, InStr(This, ";") + 1&))
    ElseIf This = "End" Then
        Exit Sub
    End If
End Sub

Function ALU(ByVal Expression As String) As String
    'Expression             식
    '>, >=, =, <=, <, != 은 논리연산 Logical Unit으로
    '+, -, *, /, ^, % 은 산술연산 Arithmetic Unit으로
        
    Dim sLeft               As String '좌항
    Dim sRight              As String '우항
    
    If InStr(Expression, ">=") Then
        sLeft = Split(Expression, ">=")(0)
        sRight = Split(Expression, ">=")(1)
        
        If Not IsNumber(sLeft) Then
            sLeft = GetValue(sLeft, "null")
            If sLeft = Null Then MsgBox "변수 불러오기 실패" '값이 불려와지지않았을경우
        End If
        If Not IsNumber(sRight) Then
            sRight = GetValue(sRight, "null")
            If sRight = Null Then MsgBox "변수 불러오기 실패"
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
            If sLeft = Null Then MsgBox "변수 불러오기 실패" '값이 불려와지지않았을경우
        End If
        If Not IsNumber(sRight) Then
            sRight = GetValue(sRight, "null")
            If sRight = Null Then MsgBox "변수 불러오기 실패"
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
            If sLeft = Null Then MsgBox "변수 불러오기 실패" '값이 불려와지지않았을경우
        End If
        If Not IsNumber(sRight) Then
            sRight = GetValue(sRight, "null")
            If sRight = Null Then MsgBox "변수 불러오기 실패"
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
            If sLeft = Null Then MsgBox "변수 불러오기 실패" '값이 불려와지지않았을경우
        End If
        If Not IsNumber(sRight) Then
            sRight = GetValue(sRight, "null")
            If sRight = Null Then MsgBox "변수 불러오기 실패"
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
            If sLeft = Null Then MsgBox "변수 불러오기 실패" '값이 불려와지지않았을경우
        End If
        If Not IsNumber(sRight) Then
            sRight = GetValue(sRight, "null")
            If sRight = Null Then MsgBox "변수 불러오기 실패"
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
            If sLeft = Null Then MsgBox "변수 불러오기 실패" '값이 불려와지지않았을경우
        End If
        If Not IsNumber(sRight) Then
            sRight = GetValue(sRight, "null")
            If sRight = Null Then MsgBox "변수 불러오기 실패"
        End If
        
        If Val(sLeft) = Val(sRight) Then
            ALU = True
        Else
            ALU = False
        End If
    '-------------------------------------------------- 여기까지 Logic Unit
    '-------------------------------------------------- 여기부턴 Arithmetic Unit
    ElseIf InStr(Expression, "+") Then
        sLeft = Split(Expression, "+")(0)
        sRight = Split(Expression, "+")(1)
        
        If Not IsNumber(sLeft) Then
            sLeft = GetValue(sLeft, "null")
            If sLeft = Null Then MsgBox "변수 불러오기 실패" '값이 불려와지지않았을경우
        End If
        If Not IsNumber(sRight) Then
            sRight = GetValue(sRight, "null")
            If sRight = Null Then MsgBox "변수 불러오기 실패"
        End If
        
        ALU = Val(sLeft) + Val(sRight)
    ElseIf InStr(Expression, "-") Then
        sLeft = Split(Expression, "-")(0)
        sRight = Split(Expression, "-")(1)
        
        If Not IsNumber(sLeft) Then
            sLeft = GetValue(sLeft, "null")
            If sLeft = Null Then MsgBox "변수 불러오기 실패" '값이 불려와지지않았을경우
        End If
        If Not IsNumber(sRight) Then
            sRight = GetValue(sRight, "null")
            If sRight = Null Then MsgBox "변수 불러오기 실패"
        End If
        
        ALU = Val(sLeft) - Val(sRight)
    ElseIf InStr(Expression, "*") Then
        sLeft = Split(Expression, "*")(0)
        sRight = Split(Expression, "*")(1)
        
        If Not IsNumber(sLeft) Then
            sLeft = GetValue(sLeft, "null")
            If sLeft = Null Then MsgBox "변수 불러오기 실패" '값이 불려와지지않았을경우
        End If
        If Not IsNumber(sRight) Then
            sRight = GetValue(sRight, "null")
            If sRight = Null Then MsgBox "변수 불러오기 실패"
        End If
        
        ALU = Val(sLeft) * Val(sRight)
    ElseIf InStr(Expression, "/") Then
        sLeft = Split(Expression, "/")(0)
        sRight = Split(Expression, "/")(1)
        
        If Not IsNumber(sLeft) Then
            sLeft = GetValue(sLeft, "null")
            If sLeft = Null Then MsgBox "변수 불러오기 실패" '값이 불려와지지않았을경우
        End If
        If Not IsNumber(sRight) Then
            sRight = GetValue(sRight, "null")
            If sRight = Null Then MsgBox "변수 불러오기 실패"
        End If
        
        ALU = Val(sLeft) / Val(sRight)
    ElseIf InStr(Expression, "^") Then
        sLeft = Split(Expression, "^")(0)
        sRight = Split(Expression, "^")(1)
        
        If Not IsNumber(sLeft) Then
            sLeft = GetValue(sLeft, "null")
            If sLeft = Null Then MsgBox "변수 불러오기 실패" '값이 불려와지지않았을경우
        End If
        If Not IsNumber(sRight) Then
            sRight = GetValue(sRight, "null")
            If sRight = Null Then MsgBox "변수 불러오기 실패"
        End If
        
        ALU = Val(sLeft) ^ Val(sRight)
    ElseIf InStr(Expression, "%") Then
        sLeft = Split(Expression, "%")(0)
        sRight = Split(Expression, "%")(1)
        
        If Not IsNumber(sLeft) Then
            sLeft = GetValue(sLeft, "null")
            If sLeft = Null Then MsgBox "변수 불러오기 실패" '값이 불려와지지않았을경우
        End If
        If Not IsNumber(sRight) Then
            sRight = GetValue(sRight, "null")
            If sRight = Null Then MsgBox "변수 불러오기 실패"
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
    
    '가장우선순위는 역시 괄호지!^^
    
    Do Until InStr(sRight, ")") = 0
        LeftPos = InStrRev(sRight, "(")
        RightPos = InStr(sRight, ")")
        sRight = Left$(sRight, LeftPos - 1&) & ALU(Mid$(sRight, LeftPos + 1&, RightPos - LeftPos - 1&)) & Right$(sRight, Len(sRight) - RightPos)
        DoEvents
    Loop
    
    Do Until InStr(sRight, "^") = 0   ' ^이 없어질때까지 계속 치환작업
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
        
        Do Until InStr(sRight, "*") = 0   '*이 없어질때까지 계속 치환작업
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
        
        Do Until InStr(sRight, "/") = 0   ' /이 없어질때까지 계속 치환작업
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
        
        Do Until InStr(sRight, "%") = 0   ' %이 없어질때까지 계속 치환작업
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
        
        Do Until InStr(sRight, "+") = 0   ' +이 없어질때까지 계속 치환작업
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
        
        Do Until InStr(sRight, "-") = 0   ' -이 없어질때까지 계속 치환작업
            '음수이여서그런걸수도잇으니..
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
