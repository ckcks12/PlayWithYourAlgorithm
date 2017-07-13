VERSION 5.00
Begin VB.Form frm_Size 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "취소"
      Height          =   255
      Left            =   8040
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "확인"
      Height          =   255
      Left            =   6480
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
   End
   Begin VB.Label Label1 
      Caption         =   "현재 값 :  "
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frm_Size"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public a As Long

Private Sub Command1_Click()
    frm_Main.WorkSheet1.SetSize HScroll1.Value
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    HScroll1.Value = a
    Label1.Caption = "현재 값 : " & HScroll1.Value
End Sub

Private Sub HScroll1_Change()
    Label1.Caption = "현재 값 : " & HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub
