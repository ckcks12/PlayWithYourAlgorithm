Attribute VB_Name = "mod_Make"
Option Explicit

Public Const TEST As Boolean = True '�����

'������(����)�� ���õ� ���
'wyodu@naver.com

'----------------------------------Ű����
'=              ������ �������� ����
'+              ���׿� ������ ����
'-              ���׿� ������ ��
'*              ����
'/              ����
'^              ������ ���׸�ŭ �ŵ�����
'print A        A�μ� (��������)
'goto 2         2������ ������ �̵�
'if A b; c;     A�� ���ǹ��˻� -  ==���� !=�ٸ��� > >= < =< OR AND ������� A�� ���̸� b �����̸� c

'���콺����Ʈ������
Public Type MPoint
    X As Long
    Y As Long
End Type

'Left Top Width Height�� ����
Public Type LTWH
    L As Integer
    T As Integer
    W As Integer
    H As Integer
End Type


Function MakeEXE(ByVal sPath As String, ByVal sDLLPath As String, ByVal PB As PropertyBag) As Boolean
    'sPath              ������ ������ �ּ�
    'sDLLPath           DLL������ �ּ�
    'PB                 �����Ͱ� ����ִ� PropertyBag
On Error GoTo ERRTABLE
    Dim LOFPoint               As Long
    Dim varTemp                 As Variant
    Dim iFF                     As Integer: iFF = FreeFile
    
    varTemp = PB.Contents 'PB.Contents�� �սǵ����ʵ��� �踮��Ʈ�� ������ ������
    FileCopy sDLLPath, sPath
    
    Open sPath For Binary As iFF
        LOFPoint = LOF(iFF)
        Seek #iFF, LOFPoint
        Put #iFF, , varTemp 'varTemp�����͸� �������
        Put #iFF, , LOFPoint '�ٽ� LOF�����ͷ� ������
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


