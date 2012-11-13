VERSION 5.00
Begin VB.Form wheel 
   Caption         =   "Form1"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   Picture         =   "wheel.frx":0000
   ScaleHeight     =   5970
   ScaleWidth      =   8850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdnext 
      BackColor       =   &H0000C000&
      Caption         =   "Go to Showcases!!"
      BeginProperty Font 
         Name            =   "@Batang"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   2775
   End
   Begin VB.CommandButton cmdspin 
      BackColor       =   &H000000C0&
      Caption         =   "Spin the Wheel!!"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   1695
   End
End
Attribute VB_Name = "wheel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdnext_Click()
'this button allows the user to continue to the next form
Showcase.Show
wheel.Hide
End Sub

Private Sub cmdspin_Click()

Dim Value As Integer
Dim Spin(1 To 10) As Single
Dim Winnings(1 To 10) As Single
Dim Ctr As Integer


Open App.Path & "\spin.txt" For Input As #1
    Ctr = 0
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Spin(Ctr), Winnings(Ctr)
Loop

Close #1
Value = InputBox("Enter a Number between 1 and 9: 1 being a slow spin and 9 being a fast spin", "How fast do you want it to spin?")
Dim Found As Boolean
Dim Pos As Integer

Found = False
Do While Found = False And Pos < Ctr
    Pos = Pos + 1
    If Value = Spin(Pos) Then
        Found = True
    End If
Loop

If Found = True Then
    MsgBox (WholeName) & " Congratulations, you have won " & FormatCurrency(Winnings(Pos)), , "WINNER"
End If
    Runningtotal = Runningtotal + Winnings(Pos)

    



End Sub
