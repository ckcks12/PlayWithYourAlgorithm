VERSION 5.00
Begin VB.UserControl WorkSheet 
   ClientHeight    =   4980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11415
   ScaleHeight     =   4980
   ScaleWidth      =   11415
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   0
      ScaleHeight     =   4815
      ScaleWidth      =   11415
      TabIndex        =   1
      Top             =   0
      Width           =   11415
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FillColor       =   &H80000006&
      ForeColor       =   &H80000006&
      Height          =   135
      Left            =   5520
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   0
      Top             =   4800
      Width           =   135
   End
End
Attribute VB_Name = "WorkSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim p3MouseDown As Boolean
Dim p3MousePoint As MPoint

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    p3MouseDown = True
    p3MousePoint.X = X
    p3MousePoint.Y = Y
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If p3MouseDown Then
 
        Picture2.height = Picture2.height + (Y - p3MousePoint.Y)
    End If
    Picture3.Top = Picture2.height
 
    
    UserControl.Width = Picture2.Width + 250
    UserControl.height = Picture2.height + 250
End Sub

Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    p3MouseDown = False
End Sub
 
Public Sub SetSize(ByVal height As Long)
    Picture2.height = height
    Picture3.Top = Picture2.height
 
    
    UserControl.Width = Picture2.Width + 250
    UserControl.height = Picture2.height + 250
End Sub


