VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Quit"
      Height          =   975
      Left            =   5280
      TabIndex        =   7
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   1095
      Left            =   5280
      TabIndex        =   6
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Your Salary Increase"
      Height          =   735
      Left            =   1800
      TabIndex        =   5
      Top             =   2760
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   840
      ScaleHeight     =   1395
      ScaleWidth      =   3915
      TabIndex        =   4
      Top             =   3600
      Width           =   3975
   End
   Begin VB.TextBox txtRank 
      Height          =   975
      Left            =   2880
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtSalaryBox 
      Height          =   975
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblRankinput 
      Caption         =   "Input Integer Rank"
      Height          =   855
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblSalaryinput 
      Caption         =   "Input Salary"
      Height          =   855
      Left            =   720
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim Sal As Single
    Dim Rank As Integer
    Dim Message As String
    

Private Sub Command1_Click()
    Rank = Rank / 2
    Sal = Sal / Rank
    Picture1.Print "Your salary increase is equal to"; (Sal)
End Sub

Private Sub Command2_Click()
    Picture1.Cls
    Rank = 0
    Sal = 0
End Sub

Private Sub Command3_Click()
    End
End Sub

Private Sub txtRank_Change()
    Let Rank = txtRank.Text
    If Rank <= 4 And Rank >= 1 Then
    Else
        Message = "Your rank needs to be between the interval of 1-4"
        MsgBox Message, , "Error"
End If
    
    
End Sub

Private Sub txtSalaryBox_Change()
    Let Sal = txtSalaryBox.Text
End Sub
