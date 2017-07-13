VERSION 5.00
Begin VB.Form frm_Play 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frm_Play.frx":0000
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   2160
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   1320
      ScaleHeight     =   675
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   960
      Width           =   2175
   End
End
Attribute VB_Name = "frm_Play"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Ready
End Sub

Private Sub Form_Load()

End Sub
