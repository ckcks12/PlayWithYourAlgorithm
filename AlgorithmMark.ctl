VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   675
   ScaleHeight     =   675
   ScaleWidth      =   675
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1740
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   0
      Picture         =   "AlgorithmMark.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   675
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public MyType As Long
'0      Terminal
'1      Process
'2      Decision
'3      Read
'4      Print
'5      Declare
'6      Connector

Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event DbClick()

 Sub SetCode(ByVal STR As String)
    Label1.Caption = STR
End Sub

 Function GetCode() As String
    GetCode = Label1.Caption
End Function

 Sub SetIndex(ByVal Index As Integer)
    Image1.Tag = Index
End Sub

 Function GetIndex() As Integer
    GetIndex = Image1.Tag
End Function

Sub Change(ByVal t As Long)
    MyType = t
    With Image1
    Select Case t
        Case 1
            .Picture = LoadPicture(App.Path & "\img\terminal.jpg")
            .Width = 2250: .height = 690
            'Label1.Left = 240: Label1.Top = 420
        Case 2
            .Picture = LoadPicture(App.Path & "\img\process.jpg")
            Label1.Visible = True
            .Width = 2250: .height = 690
            'Label1.Left = 240: Label1.Top = 420
        Case 3
            .Picture = LoadPicture(App.Path & "\img\decision.jpg")
            Label1.Visible = True
            .Width = 2250: .height = 690
            'Label1.Left = 225: Label1.Top = 420
        Case 4
            .Picture = LoadPicture(App.Path & "\img\read.jpg")
            Label1.Visible = True
            .Width = 2250: .height = 690
            'Label1.Left = 240:  Label1.Top = 600
        Case 5
            .Picture = LoadPicture(App.Path & "\img\print.jpg")
            Label1.Visible = True
            .Width = 2250: .height = 690
            'Label1.Left = 240:  Label1.Top = 240
        Case 6
            .Picture = LoadPicture(App.Path & "\img\declare.jpg")
            Label1.Visible = True
            .Width = 2250: .height = 690
            'Label1.Left = 420:  Label1.Top = 240
        Case 7
            .Picture = LoadPicture(App.Path & "\img\connector.jpg")
            UserControl.Width = 675
            UserControl.height = 675
            .Width = 675
            .height = 675
            Label1.Visible = False
        Case 8
            .Picture = LoadPicture(App.Path & "\img\for.jpg")
            Label1.Visible = True
            .Width = 2250: .height = 690
        Case 9
            .Picture = LoadPicture(App.Path & "\img\back.jpg")
            .Width = 2250: .height = 690
        Case Else
    End Select
    End With
End Sub

Sub Selected(ByVal sure As Boolean)
    If sure Then
        With Image1
        Select Case MyType
            Case 1
                .Picture = LoadPicture(App.Path & "\img\terminal_.jpg")
                '.Width = 2250: .Height = 690
            Case 2
                .Picture = LoadPicture(App.Path & "\img\process_.jpg")
                '.Width = 2250: .Height = 690
            Case 3
                .Picture = LoadPicture(App.Path & "\img\decision_.jpg")
            Case 4
                .Picture = LoadPicture(App.Path & "\img\read_.jpg")
            Case 5
                .Picture = LoadPicture(App.Path & "\img\print_.jpg")
            Case 6
                .Picture = LoadPicture(App.Path & "\img\declare_.jpg")
            Case 7
                .Picture = LoadPicture(App.Path & "\img\connector_.jpg")
                UserControl.Width = 675
            UserControl.height = 675
            .Width = 675
            .height = 675
            Label1.Visible = False
            Case 8
                .Picture = LoadPicture(App.Path & "\img\for_.jpg")
            Case 9
                .Picture = LoadPicture(App.Path & "\img\back_.jpg")
            Case Else
        End Select
        End With
    Else
        With Image1
        Select Case MyType
            Case 1
                .Picture = LoadPicture(App.Path & "\img\terminal.jpg")
                '.Width = 2250: .Height = 690
            Case 2
                .Picture = LoadPicture(App.Path & "\img\process.jpg")
                '.Width = 2250: .Height = 690
            Case 3
                .Picture = LoadPicture(App.Path & "\img\decision.jpg")
            Case 4
                .Picture = LoadPicture(App.Path & "\img\read.jpg")
            Case 5
                .Picture = LoadPicture(App.Path & "\img\print.jpg")
            Case 6
                .Picture = LoadPicture(App.Path & "\img\declare.jpg")
            Case 7
                .Picture = LoadPicture(App.Path & "\img\connector.jpg")
                UserControl.Width = 675
            UserControl.height = 675
            .Width = 675
            .height = 675
            Label1.Visible = False
            Case 8
                .Picture = LoadPicture(App.Path & "\img\for.jpg")
            Case 9
                .Picture = LoadPicture(App.Path & "\img\back.jpg")
            Case Else
        End Select
        End With
    End If
End Sub

Private Sub Image1_Click()
    RaiseEvent Click
End Sub

Private Sub Image1_DblClick()
    RaiseEvent DbClick
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Label1_Click()
    RaiseEvent Click
End Sub

Private Sub Label1_DblClick()
    RaiseEvent DbClick
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DbClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
    If MyType = 7 Then Exit Sub
    With UserControl
        .Width = 2250: .height = 690
    End With
End Sub
