Attribute VB_Name = "mod_Make"
Option Explicit

Public Const TEST As Boolean = True '디버그

'컴파일(생성)에 관련된 모듈
'wyodu@naver.com

'----------------------------------키워드
'=              우항을 좌항으로 대입
'+              우항에 좌항을 더함
'-              우항에 좌항을 뺌
'*              곱함
'/              나눔
'^              우항을 좌항만큼 거듭제곱
'print A        A인쇄 (결과값출력)
'goto 2         2번으로 포인터 이동
'if A b; c;     A식 조건문검사 -  ==같다 !=다르다 > >= < =< OR AND 연산까지 A가 참이면 b 거짓이면 c

'마우스포인트형변수
Public Type MPoint
    X As Long
    Y As Long
End Type

'Left Top Width Height형 변수
Public Type LTWH
    L As Integer
    T As Integer
    W As Integer
    H As Integer
End Type


Function MakeEXE(ByVal sPath As String, ByVal sDLLPath As String, ByVal PB As PropertyBag) As Boolean
    'sPath              생성할 파일의 주소
    'sDLLPath           DLL파일의 주소
    'PB                 데이터가 들어있는 PropertyBag
On Error GoTo ERRTABLE
    Dim LOFPoint               As Long
    Dim varTemp                 As Variant
    Dim iFF                     As Integer: iFF = FreeFile
    
    varTemp = PB.Contents 'PB.Contents를 손실되지않도록 배리언트형 변수에 담은후
    FileCopy sDLLPath, sPath
    
    Open sPath For Binary As iFF
        LOFPoint = LOF(iFF)
        Seek #iFF, LOFPoint
        Put #iFF, , varTemp 'varTemp데이터를 써넣은후
        Put #iFF, , LOFPoint '다시 LOF포인터로 마무리
    Close #iFF
    
    MakeEXE = True
    Exit Function
    
ERRTABLE:
    #If TEST Then
        MsgBox Err.Number & vbCrLf & Err.Description
    #End If
    MakeEXE = False
End Function

Function SplitEX(ByVal STR As String, ByVal sStart As String, ByVal sEnd As String) As String
    SplitEX = Split(Split(STR, sStart)(1), sEnd)(0)
End Function


