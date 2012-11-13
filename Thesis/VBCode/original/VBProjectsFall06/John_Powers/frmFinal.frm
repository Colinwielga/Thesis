VERSION 5.00
Begin VB.Form frmFinal 
   Caption         =   "Final Form"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRight 
      Caption         =   "Both are wrong"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "2"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "1"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lbl2 
      Caption         =   "2. Multipied the number by 2"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "1. Divided the number by 2"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblFinal 
      Caption         =   "What did the last program do?"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmFinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd1_Click()
    MsgBox "Sorry, but you're wrong.", , "Incorrect"
    frmFinal.Hide
    frmStart.Show

End Sub

Private Sub cmd2_Click()
    MsgBox "Sorry, but you're wrong.", , "Incorrect"
    frmFinal.Hide
    frmStart.Show

End Sub

Private Sub cmdBack_Click()
    frmNumber.Show
    frmFinal.Hide

End Sub


Private Sub cmdRight_Click()
    MsgBox "CORRECT!", , "Correct"
    frmFinal.Hide
    frmBonus.Show

End Sub
