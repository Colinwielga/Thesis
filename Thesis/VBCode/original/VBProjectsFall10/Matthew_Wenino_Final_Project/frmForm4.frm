VERSION 5.00
Begin VB.Form frmForm4 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form4"
   ClientHeight    =   5475
   ClientLeft      =   6510
   ClientTop       =   4830
   ClientWidth     =   4620
   BeginProperty Font 
      Name            =   "Lucida Bright"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   Picture         =   "frmForm4.frx":0000
   ScaleHeight     =   5475
   ScaleWidth      =   4620
   Begin VB.Timer tmrTest 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3000
      Top             =   120
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop Timer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      TabIndex        =   1
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Timer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   4080
      Width           =   1335
   End
End
Attribute VB_Name = "frmForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private secCount As Integer

Private Sub cmdStart_Click()
secCount = 6
Cls
tmrTest.Enabled = True
End Sub


Private Sub cmdStop_Click()
tmrTest.Enabled = False
End Sub


Private Sub tmrTest_Timer()
secCount = secCount - 1
Print "You have " & secCount; " second(s) until shutdown."
If secCount = 0 Then
    tmrTest.Enabled = False
    MsgBox "Thank you for using the Virtual Match-Maker 2.0! Goodbye!"
    End
End If
End Sub

