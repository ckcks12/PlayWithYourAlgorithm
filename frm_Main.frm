VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Play! with your Algorithm! "
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   15255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   15255
   Begin MSComDlg.CommonDialog CD 
      Left            =   1800
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "fwc����|*.fwc"
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   0
      TabIndex        =   19
      Top             =   5400
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "NEXT"
         Height          =   615
         Left            =   2640
         TabIndex        =   24
         Top             =   840
         Width           =   1695
      End
      Begin VB.ListBox listValue 
         Height          =   2205
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   4215
      End
      Begin VB.Label Label3 
         Caption         =   "�νĵ� �帧�� ��ü�� ����"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "�νĵ� ������ ��ü�� ����"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   3975
      End
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   8685
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "fdfdf"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   17
      Top             =   9060
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   53
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox back 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   2280
      Picture         =   "frm_Main.frx":0000
      ScaleHeight     =   1050
      ScaleWidth      =   2250
      TabIndex        =   14
      Top             =   3465
      Width           =   2250
   End
   Begin VB.PictureBox connector 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   840
      Picture         =   "frm_Main.frx":0C74
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   13
      Top             =   4665
      Width           =   675
   End
   Begin VB.PictureBox ffor 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   0
      Picture         =   "frm_Main.frx":1427
      ScaleHeight     =   1050
      ScaleWidth      =   2250
      TabIndex        =   12
      Top             =   3465
      Width           =   2250
   End
   Begin VB.PictureBox ddeclare 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   2280
      Picture         =   "frm_Main.frx":20DB
      ScaleHeight     =   1050
      ScaleWidth      =   2250
      TabIndex        =   11
      Top             =   2385
      Width           =   2250
   End
   Begin VB.PictureBox pprint 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   0
      Picture         =   "frm_Main.frx":2D6A
      ScaleHeight     =   1050
      ScaleWidth      =   2250
      TabIndex        =   10
      Top             =   2385
      Width           =   2250
   End
   Begin VB.PictureBox read 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   2280
      Picture         =   "frm_Main.frx":3842
      ScaleHeight     =   1050
      ScaleWidth      =   2250
      TabIndex        =   9
      Top             =   1305
      Width           =   2250
   End
   Begin VB.PictureBox decision 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   0
      Picture         =   "frm_Main.frx":4307
      ScaleHeight     =   1050
      ScaleWidth      =   2250
      TabIndex        =   8
      Top             =   1305
      Width           =   2250
   End
   Begin VB.PictureBox process 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   2280
      Picture         =   "frm_Main.frx":53F3
      ScaleHeight     =   1050
      ScaleWidth      =   2250
      TabIndex        =   7
      Top             =   225
      Width           =   2250
   End
   Begin VB.PictureBox terminal 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   0
      Picture         =   "frm_Main.frx":5BEC
      ScaleHeight     =   1050
      ScaleWidth      =   2250
      TabIndex        =   6
      Top             =   225
      Width           =   2250
   End
   Begin VB.PictureBox Picture1 
      Height          =   8655
      Left            =   4560
      ScaleHeight     =   8595
      ScaleWidth      =   10395
      TabIndex        =   1
      Top             =   0
      Width           =   10455
      Begin PlayWithYourAlgorithm.Lines Lines 
         Height          =   135
         Index           =   0
         Left            =   1080
         TabIndex        =   16
         Top             =   5640
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   2990
      End
      Begin PlayWithYourAlgorithm.UserControl1 OBJ 
         Height          =   690
         Index           =   0
         Left            =   2040
         TabIndex        =   5
         Top             =   5760
         Visible         =   0   'False
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   1217
      End
      Begin VB.Timer timer_lines 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   6360
         Top             =   7200
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   6120
         Top             =   6240
      End
      Begin PlayWithYourAlgorithm.WorkSheet WorkSheet1 
         Height          =   4935
         Left            =   -840
         TabIndex        =   2
         Top             =   0
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   8705
      End
      Begin VB.Label b 
         Caption         =   "Label1"
         Height          =   255
         Left            =   10560
         TabIndex        =   4
         Top             =   5880
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label a 
         Caption         =   "Label1"
         Height          =   255
         Left            =   10080
         TabIndex        =   3
         Top             =   6960
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   8655
      LargeChange     =   700
      Left            =   15000
      Max             =   0
      SmallChange     =   700
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   2640
      Picture         =   "frm_Main.frx":67CA
      Top             =   4920
      Width           =   1500
   End
   Begin VB.Label Label2 
      Caption         =   "������ ���� ����"
      Height          =   225
      Left            =   1455
      TabIndex        =   15
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   2280
      TabIndex        =   23
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Menu mFile 
      Caption         =   "����(&F)"
      Index           =   0
      Begin VB.Menu mOpen 
         Caption         =   "����(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mSave 
         Caption         =   "����(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mAnotherSave 
         Caption         =   "�ٸ��̸���������(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mExit 
         Caption         =   "������(&X)"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mEdit 
      Caption         =   "����(&E)"
      Begin VB.Menu mSize 
         Caption         =   "����ũ��(&Z)"
         Shortcut        =   ^Z
      End
   End
   Begin VB.Menu mCode 
      Caption         =   "�ڵ�(&C)"
      Begin VB.Menu mVB 
         Caption         =   "Visual Basic"
      End
      Begin VB.Menu mC 
         Caption         =   "C"
      End
   End
   Begin VB.Menu mExecute 
      Caption         =   "����(&U)"
      Begin VB.Menu mStep 
         Caption         =   "�ܰ躰����(&T)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mMake 
         Caption         =   "EXE�����(&M)"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mAbout 
         Caption         =   "������"
      End
      Begin VB.Menu mHelper 
         Caption         =   "����"
      End
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const TEST As Boolean = True
Const ScrollSense As Integer = 300 '������ ��ü �ڵ����� grid�� ���߱�

Dim ObjClick(9) As Boolean
Dim MP As MPoint
Dim Clicked As Boolean
Dim Connection As String
Dim LinesClicked As Boolean
Dim ObjTop As New Collection
Dim LinesTop As New Collection



Private Sub Commadnd1_Click()
    Dim PB As New PropertyBag
    Dim sTemp() As String, sTemp2() As String, i As Long, ValueCount As Long, sTemp3 As String
    
    PB.WriteProperty "caption", "�ȳ�"
           
    sTemp2 = Split(Replace$(Text1.Text, vbCrLf, vbNullString), "||")
    
    For i = 0 To UBound(sTemp2) '�ҽ��� �ϳ��ϳ��� �и��Ͽ� �����
        PB.WriteProperty i, Trim$(sTemp2(i))
        DoEvents
    Next i
    
    PB.WriteProperty "count", UBound(sTemp2)
    
    MsgBox MakeEXE(App.Path & "\Play.exe", App.Path & "\play.dll", PB)
End Sub

Private Sub HScroll1_Change()
    WorkSheet1.Left = -1 * (HScroll1.Value)
End Sub

Private Sub HScroll1_Scroll()
    Call HScroll1_Change
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    '���� Start�� ã�´�
    Dim i As Integer, j As Integer, k As Integer
    
    For i = 1 To OBJ.UBound
        If OBJ(i).MyType = 1 And OBJ(i).GetCode = "Start" Then Exit For
    Next i
    
    'ã������ ���� �帧������ ã�´�!
    
    For j = 1 To Lines.UBound
        If Val(Split(Lines(j).Tag, " ")(0)) = i Then Exit For
    Next j
    
    'ã������ �״����� ����?
    i = Val(Split(Lines(j).Tag, " ")(1))
    
    Do Until OBJ(i).GetCode = "End"
        For j = 1 To Lines.UBound
            If Val(Split(Lines(j).Tag, " ")(0)) = i Then Exit For
            DoEvents
        Next
        i = Val(Split(Lines(j).Tag, " ")(1))
        k = j
        DoEvents
        MsgBox "������"
        OBJ(i).Selected True
    Loop
End Sub

Private Sub Form_Load()
    '----��������ü�������� img�������� �������� ��������
    On Error Resume Next
    terminal.Picture = LoadPicture(App.Path & "\img\terminal.jpg")
    process.Picture = LoadPicture(App.Path & "\img\process.jpg")
    decision.Picture = LoadPicture(App.Path & "\img\decision.jpg")
    read.Picture = LoadPicture(App.Path & "\img\read.jpg")
    pprint.Picture = LoadPicture(App.Path & "\img\print.jpg")
    ddeclare.Picture = LoadPicture(App.Path & "\img\declare.jpg")
    ffor.Picture = LoadPicture(App.Path & "\img\for.jpg")
    back.Picture = LoadPicture(App.Path & "\img\back.jpg")
    connector.Picture = LoadPicture(App.Path & "\img\connector.jpg")
    '----
    
End Sub

Private Sub Image1_Click()
    If Connection = "" Then
        timer_lines.Enabled = True
        LinesClicked = True
    Else
        timer_lines.Enabled = False
        LinesClicked = True
        timer_lines.Enabled = True
    End If
        
End Sub

Private Sub Label4_Click()
    Call Image1_Click
End Sub

Private Sub Lines_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Unload Lines(Index): Exit Sub
End Sub

Private Sub mOpen_Click()
    '����
    Dim i As Integer, PB As New PropertyBag, b() As Byte, var As Variant, pos As Long
    
    CD.DefaultExt = App.Path
    CD.CancelError = True: On Error GoTo exit001
    CD.ShowOpen

            i = FreeFile
            Open CD.FileName For Binary As #i
                Get #i, LOF(i) - 3, pos
                Seek #i, pos
                Get #i, , var
            Close #i
            b = var
            PB.Contents = b
            '---������ ��� ����� �۾�
            For i = 1 To OBJ.UBound
                Unload OBJ(i)
                DoEvents
            Next i
            For i = 1 To Lines.UBound
                Unload Lines(i)
                DoEvents
            Next i
            '---�ҷ������۾�
            For i = 1 To Val(PB.ReadProperty("obj"))
                On Error Resume Next
                Load OBJ(i)
                OBJ(i).SetCode PB.ReadProperty("obj" & i & "code")
                OBJ(i).Change PB.ReadProperty("obj" & i & "type")
                OBJ(i).Left = PB.ReadProperty("obj" & i & "left")
                OBJ(i).top = PB.ReadProperty("obj" & i & "top")
                ObjTopAdd OBJ(i).top, CStr(i)
                OBJ(i).Visible = True
                OBJ(i).ZOrder 0
                DoEvents
            Next i
            For i = 1 To Val(PB.ReadProperty("lines"))
                Load Lines(i)
                Lines(i).Tag = PB.ReadProperty("lines" & i & "tag")
                Lines(i).SetLine PB.ReadProperty("lines" & i & "type"), PB.ReadProperty("lines" & i & "width"), PB.ReadProperty("lines" & i & "height")
                If Not PB.ReadProperty("lines" & i & "yesno") = 0 Then
                    If PB.ReadProperty("lines" & i & "yesno") = 1 Then Lines(i).SetYESNO True
                    If PB.ReadProperty("lines" & i & "yesno") = 2 Then Lines(i).SetYESNO False
                End If
                Lines(i).Left = PB.ReadProperty("lines" & i & "left")
                Lines(i).top = PB.ReadProperty("lines" & i & "top")
                Lines(i).Visible = True
                LinesTopADD Lines(i).top, CStr(i)
                Lines(i).ZOrder 0
                DoEvents
            Next i
            WorkSheet1.height = PB.ReadProperty("worksheet")
            Log "������ ���������� �ҷ��Խ��ϴ�"

exit001:
End Sub

Private Sub mSave_Click()
    '�����ϱ�
    Dim i As Integer, PB As New PropertyBag, var As Variant, pos As Long
    
    CD.DefaultExt = App.Path
    CD.CancelError = True: On Error GoTo exit001
    CD.ShowSave
    If Len(CD.FileName) > 2 Then
        For i = 1 To OBJ.UBound
            PB.WriteProperty "obj" & i & "code", OBJ(i).GetCode
            PB.WriteProperty "obj" & i & "type", OBJ(i).MyType
            PB.WriteProperty "obj" & i & "left", OBJ(i).Left
            PB.WriteProperty "obj" & i & "top", OBJ(i).top
            DoEvents
        Next i
        PB.WriteProperty "obj", OBJ.UBound
        
        For i = 1 To Lines.UBound
            PB.WriteProperty "lines" & i & "left", Lines(i).Left
            PB.WriteProperty "lines" & i & "top", Lines(i).top
            PB.WriteProperty "lines" & i & "tag", Lines(i).Tag
            PB.WriteProperty "lines" & i & "type", Lines(i).MyType
            PB.WriteProperty "lines" & i & "height", Lines(i).height
            PB.WriteProperty "lines" & i & "width", Lines(i).Width
            PB.WriteProperty "lines" & i & "yesno", Lines(i).yesno
            DoEvents
        Next i
        PB.WriteProperty "lines", Lines.UBound
        PB.WriteProperty "worksheet", WorkSheet1.height
        i = FreeFile
        Open CD.FileName For Output As #i
            Print #i, "D"
        Close #i
        var = PB.Contents
        Open CD.FileName For Binary As #i
            pos = LOF(i)
            Seek #i, pos
            Put #i, , var
            Put #i, , pos
        Close #i
        MsgBox "����Ǿ����ϴ�", vbSystemModal, "����"
        Log Time & "  ����Ǿ����ϴ�"
    End If
exit001:
End Sub

Private Sub mSize_Click()
    Load frm_Size
    frm_Size.a = WorkSheet1.height
    frm_Size.HScroll1.Value = WorkSheet1.height
    frm_Size.Show
End Sub

Private Sub OBJ_Click(Index As Integer)
    If LinesClicked Then '�帧���׸������ǿ�?
        OBJ(Index).Selected (True) '�ϴ� �� ��ĥ�����
        Connection = Connection & Index & " "
    End If
End Sub

Private Sub OBJ_DbClick(Index As Integer)
    Dim Temp$, i As Integer
    Select Case OBJ(Index).MyType
        Case 1 'terminal
            Temp = InputBox("Terminal �����Դϴ�" & vbNewLine & "Start �Ǵ� End�� �Է����ֽʽÿ�" & vbNewLine & _
                             "Start�� End�� �ϳ����� �߰��� �˴ϴ�" & vbCrLf & "��ҹ��ڴ� ���е˴ϴ�")
            For i = 1 To OBJ.UBound
                If OBJ(i).MyType = 1 And OBJ(i).GetCode = "Start" And Temp = "Start" Then
                    MsgBox "�̹� Start�� �����մϴ�!", vbCritical, "����"
                    Exit Sub
                ElseIf OBJ(i).MyType = 1 And OBJ(i).GetCode = "End" And Temp = "End" Then
                    MsgBox "�̹� End�� �����մϴ�!", vbCritical, "����"
                    Exit Sub
                End If
            Next i
            If Temp = "Start" Then OBJ(Index).SetCode "Start"
            If Temp = "End" Then OBJ(Index).SetCode "End"
        
        Case 2         'Process
            Temp = InputBox("Process �����Դϴ�" & vbNewLine & _
                            "���� �Է��Ͽ� �ֽʽÿ�" & vbNewLine & _
                            "�����ڴ� ����(+), ����(-), ����(*), ������(/), ����(^), ������(%)�� �ֽ��ϴ�" & vbNewLine & _
                            "���� ���ʿ� �ִ� ��ȣ�� �켱���� ���� ����˴ϴ�")
            OBJ(Index).SetCode Replace$(Temp, " ", vbNullString)
        
        Case 3 'Decision
            Temp = InputBox("Decision �����Դϴ�" & vbNewLine & _
                            "���� �Է��Ͽ� �ֽʽÿ�" & vbNewLine & _
                            "�������ڴ� ũ��(>), ũ�ų�����(>=), ����(==), �۰ų�����(<=), �۴�(<), �ٸ���(!=)�� �ֽ��ϴ�" & vbNewLine & _
                            "���� ���ʿ� �ִ� ��ȣ�� �켱���� ���� ����˴ϴ�")
            OBJ(Index).SetCode Replace$(Temp, " ", vbNullString)
        
        Case 4 'read
            Temp = InputBox("Read �����Դϴ�" & vbNewLine & _
                             "�Է¹��� ���� ����� �������� �Է��Ͽ� �ֽʽÿ�")
            OBJ(Index).SetCode Replace$(Temp, " ", vbNullString)
        
        Case 5 'print
            Temp = InputBox("Print �����Դϴ�" & vbNewLine & _
                            "����� �������� �Է��Ͽ� �ֽʽÿ�")
            OBJ(Index).SetCode Replace$(Temp, " ", vbNullString)
        
        Case 6 'Declare
            Temp = InputBox("Declare �����Դϴ�" & vbNewLine & _
                            "������ ������ �ʱ�ȭ ���� �Է��Ͽ� �ֽʽÿ�" & vbNewLine & _
                            "ex) a��� ���� ����� a=0�� ���� �׻� �ʱⰪ�� �Է��ϼž��մϴ�" & vbNewLine & _
                            "ex) �迭��������� () �� �ƴ϶� []�� �ϼž��մϴ�" & vbNewLine & _
                            "ex) �������� ��������� �޸�(,)�� �����Ͻø� �˴ϴ�")
            OBJ(Index).SetCode Replace$(Temp, " ", vbNullString)
                            
        Case 8 'For
            Temp = InputBox("For �����Դϴ�" & vbNewLine & _
                            "�ݺ������� ���� �ӽú������ �ʱⰪ�� ���Ѱ��� �Է��Ͽ� �ֽʽÿ�" & vbNewLine & _
                            "ex) i=3,5" & vbNewLine & _
                            "ex) ������=�ʱⰪ,���Ѱ�")
            OBJ(Index).SetCode Replace$(Temp, " ", vbNullString)
    
    End Select
            
End Sub

Private Sub OBJ_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Unload OBJ(Index): Exit Sub
    Clicked = True
    MP.X = X
    MP.Y = Y
End Sub

Private Sub OBJ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Clicked And Index > 0 Then
        OBJ(Index).Left = OBJ(Index).Left + (X - MP.X) - ((X - MP.X) Mod ScrollSense)
        OBJ(Index).top = OBJ(Index).top + (Y - MP.Y) - ((Y - MP.Y) Mod ScrollSense)
    End If
End Sub

Private Sub OBJ_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ObjTop.Remove CStr(Index)
    ObjTop.Add OBJ(Index).top, CStr(Index)
    Clicked = False
End Sub

Private Sub terminal_Click()  'Terminal
    subObjClick 1 '�����߰��ϱ�
End Sub

Private Sub process_Click()  'Process
    subObjClick 2
End Sub

Private Sub decision_Click()  'Decision
    subObjClick 3
End Sub

Private Sub pprint_Click()  'Print
    subObjClick 5
End Sub

Private Sub ddeclare_Click()  'Declare
    subObjClick 6
End Sub

Private Sub read_Click()  'Read
    subObjClick 4
End Sub

Private Sub ffor_Click()  'For
    subObjClick 8
End Sub

Private Sub back_Click()  'Back
    subObjClick 9
End Sub

Private Sub connector_Click()  'Connector
    subObjClick 7
End Sub
 
Private Sub terminal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sb.SimpleText = "�͹̳��� �߰��մϴ�"
End Sub
Private Sub process_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sb.SimpleText = "���μ����� �߰��մϴ�"
End Sub
Private Sub read_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sb.SimpleText = "�б⸦ �߰��մϴ�"
End Sub
Private Sub pprint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sb.SimpleText = "����� �߰��մϴ�"
End Sub
Private Sub decision_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sb.SimpleText = "�Ǻ��� �߰��մϴ�"
End Sub
Private Sub ddeclare_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sb.SimpleText = "������ʱ�ȭ�� �߰��մϴ�"
End Sub
Private Sub ffor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sb.SimpleText = "�ݺ��� �߰��մϴ�"
End Sub
Private Sub back_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sb.SimpleText = "�ݺ����� �߰��մϴ�"
End Sub
Private Sub connector_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sb.SimpleText = "�������� �߰��մϴ�"
End Sub
Private Sub image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sb.SimpleText = "�帧���� �߰��մϴ�. Ŭ���ϰ� ������ �� ��ü�� ���ʷ� Ŭ���Ͽ� �����մϴ�"
End Sub

Private Sub timer_lines_Timer()
    '���� connection ������ �˻��ϰ�
    '���⿡ ���� line�� �ҷ����� tag�� ��ְ� tooltiptext���� ��ְ�!
    If Connection Like "* * " Then
        
        '���� �պκ� �˻��ϱ�
        Dim Temp1$, INDEX1&, Temp2$, INDEX2&

        LinesClicked = False
        
        INDEX1 = CLng(Split(Connection, " ")(0))
        INDEX2 = CLng(Split(Connection, " ")(1))
        
        OBJ(INDEX1).Selected (False)
        OBJ(INDEX2).Selected (False)
        
        If INDEX1 = INDEX2 Then timer_lines.Enabled = False: Connection = vbNullString: Exit Sub
        
        With OBJ(INDEX1)
            a.Left = .Left: a.top = .top: a.height = .height: a.Width = .Width
        End With
        With OBJ(INDEX2)
            b.Left = .Left: b.top = .top: b.height = .height: b.Width = .Width
        End With
        
        If a.Left = b.Left Or a.top = b.top Then
            If OBJ(INDEX1).MyType = 3 Then 'if���̶�� yes no ���������
                If a.top = b.top Then
                    If MsgBox("�帧���� � �������� �������?", vbYesNo, "") = vbYes Then
                        Call LinesEX(a, b, 1)
                    Else
                        Call LinesEX(a, b, 2)
                    End If
                Else
                    Log "Decision(�Ǻ�) ������ ��ü�� �������̳� ���ʿ����� �帧���� ���� �� �ֽ��ϴ�"
                    Connection = vbNullString
                End If
            Else
                Call LinesEX(a, b)
            End If
        ElseIf OBJ(INDEX1).MyType = 7 Or OBJ(INDEX2).MyType = 7 Then
                Call LinesEX(a, b)
                Connection = vbNullString
        Else
            Log "�帧���� �ձ� ���ؼ��� ����, Ȥ�� ���ΰ� ���ƾ� �մϴ�"
            Connection = vbNullString
        End If
        Connection = vbNullString
        timer_lines.Enabled = False
    End If
    
End Sub

Public Function LinesEX(ByVal aa As Label, ByVal bb As Label, Optional yesno As Integer = 0)
    'a���� b�� �ߴ°Ŵ�.
    Dim i As Integer
    Do
        i = i + 1
        On Error GoTo a
        Lines(i).Visible = True
    Loop
a:
    Load Lines(i)
    
    With Lines(i)
    
        If aa.top + aa.height < bb.top Then
            .Left = aa.Left + (aa.Width / 2) - 70&
            .top = aa.top + aa.height
            .SetLine 2, , bb.top - (aa.top + aa.height)
        ElseIf aa.top < bb.top Then
            If aa.Left < bb.Left Then
                .Left = aa.Left + aa.Width
                .top = aa.top + (aa.height / 2) - 70&
                .SetLine 0, bb.Left - (aa.Left + aa.Width)
            Else
                .Left = bb.Left + bb.Width
                .top = aa.top + (aa.height / 2) - 70&
                .SetLine 1, aa.Left - (bb.Left + bb.Width)
            End If
        ElseIf bb.top + bb.height < aa.top Then
            .Left = bb.Left + (bb.Width / 2) - 70&
            .top = bb.top + bb.height
            .SetLine 3, , aa.top - (bb.top + bb.height)
        ElseIf bb.top <= aa.top Then
            If aa.Left < bb.Left Then
                .Left = aa.Left + aa.Width
                .top = bb.top + (bb.height / 2) - 70&
                .SetLine 0, bb.Left - (aa.Left + aa.Width)
            Else
                .Left = bb.Left + bb.Width
                .top = bb.top + (bb.height / 2) - 70&
                .SetLine 1, aa.Left - (bb.Left + bb.Width)
            End If
        End If
        
        LinesTopADD .top, CStr(i)
        If yesno = 1 Then .SetYESNO True
        If yesno = 2 Then .SetYESNO False
        .Visible = True
        .ZOrder 0
        .Tag = Connection
    Connection = vbNullString
    End With
End Function

Sub LinesTopADD(ByVal top As Integer, ByVal key As String)
    On Error Resume Next
    LinesTop.Remove key
    LinesTop.Add top, key
End Sub

Sub ObjTopAdd(ByVal top As Integer, ByVal key As String)
    On Error Resume Next
    ObjTop.Remove key
    ObjTop.Add top, key
End Sub

Private Sub Timer1_Timer() '��ũ�����ִ°�
    On Error Resume Next
    If WorkSheet1.height - Picture1.height > 0 Then
        VScroll1.Max = WorkSheet1.height - Picture1.height + 250
    Else
        VScroll1.Max = 0
    End If
    '�׸��� ���簹�����˷�����
    Label1 = "�νĵ� ������ ��ü�� ���� : " & OBJ.Count - 1&
    Label3 = "�νĵ� �帧�� ��ü�� ���� : " & Lines.Count - 1&
End Sub

Private Sub VScroll1_Change()
    WorkSheet1.top = -1 * VScroll1.Value
    Dim i As Integer
    For i = 1 To OBJ.UBound
        On Error Resume Next
        OBJ(i).top = ObjTop(i) - VScroll1.Value
        DoEvents
    Next i
    For i = 1 To Lines.UBound
        On Error Resume Next
        Lines(i).top = LinesTop(i) - VScroll1.Value
        DoEvents
    Next i
End Sub

Private Sub VScroll1_Scroll()
    Call VScroll1_Change
End Sub

Sub subObjClick(ByVal T As Integer) '������ ��ü ���� �����ϴ� ����
'0      lines
'1      terminal
'2      process
'3      dicision
'4      read
'5      print
'6      declare
'7      connector
'8      for
'9      back
        Dim i As Integer
        
        Do
            i = i + 1
            On Error GoTo a
            OBJ(i).Visible = True
        Loop
a:
        Load OBJ(i)
        
        With OBJ(i)
            .Left = 0: .top = 0: .ZOrder 0: .Visible = True
            .Change T
            ObjTopAdd .top, CStr(i)
        End With
End Sub

Sub GetIndexString(ByVal Index As Integer)
    Select Case Index
        Case 1: GetIndexString = "terminal"
        Case 2: GetIndexString = "process"
        Case 3: GetIndexString = "dicision"
        Case 4: GetIndexString = "read"
        Case 5: GetIndexString = "print"
        Case 6: GetIndexString = "declare"
        Case 7: GetIndexString = "connector"
        Case 8: GetIndexString = "for"
        Case 9: GetIndexString = "back"
    End Select
End Sub

Sub Log(ByVal STR As String)
    sb.SimpleText = STR
End Sub

