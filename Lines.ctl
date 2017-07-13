VERSION 5.00
Begin VB.UserControl Lines 
   ClientHeight    =   135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
   ScaleHeight     =   135
   ScaleWidth      =   1695
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Image line2 
      Height          =   135
      Left            =   0
      Picture         =   "Lines.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "Lines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public MyType As Integer
Public YESNO As Integer

Public Event Click()
Public Event DbClick()
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Sub SetLine(ByVal WhatType As Integer, Optional ByVal Width As Integer, Optional ByVal height As Integer)
    MyType = WhatType
    Select Case WhatType
        Case 0 'right
            line2.Picture = LoadPicture(App.Path & "\img\rline.jpg")
            line2.height = 135: UserControl.height = 135
            line2.Width = Width
            UserControl.Width = Width
        Case 1 'left
            line2.Picture = LoadPicture(App.Path & "\img\lline.jpg")
            line2.height = 135: UserControl.height = 135
            line2.Width = Width
            UserControl.Width = Width
        Case 2 'down
            line2.Picture = LoadPicture(App.Path & "\img\dline.jpg")
            line2.height = height: UserControl.height = height
            line2.Width = 135
            UserControl.Width = 135
        Case 3 'up
            line2.Picture = LoadPicture(App.Path & "\img\uline.jpg")
            line2.height = height: UserControl.height = height
            line2.Width = 135
            UserControl.Width = 135
        Case Else
    End Select
    
End Sub

Sub SetYESNO(ByVal surE As Boolean)
    If surE Then 'yes
        Label1.BackStyle = 1
        Label1.BackColor = vbRed
        If MyType = 1 Then Label1.Left = UserControl.Width - (Label1.Width)
        YESNO = 1
    Else
        Label1.BackStyle = 1
        Label1.BackColor = vbBlue
        If MyType = 1 Then Label1.Left = UserControl.Width - (Label1.Width)
        YESNO = 2
    End If
End Sub

Private Sub line2_Click()
    RaiseEvent Click
End Sub
 

Private Sub line2_DblClick()
    RaiseEvent DbClick
End Sub

Private Sub line2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DbClick
End Sub

Private Sub UserControl_Initialize()
    YESNO = 0
End Sub
