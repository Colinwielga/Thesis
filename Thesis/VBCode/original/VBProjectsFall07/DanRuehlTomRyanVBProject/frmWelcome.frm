VERSION 5.00
Begin VB.Form frmWelcome 
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14655
   LinkTopic       =   "Form1"
   ScaleHeight     =   15240
   ScaleWidth      =   25080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Caption         =   "Lets Play Jeopardy!"
      Height          =   1815
      Left            =   10200
      TabIndex        =   0
      Top             =   8160
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   15360
      Left            =   120
      Picture         =   "frmWelcome.frx":0000
      Top             =   120
      Width           =   19200
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
Dim pos As Integer
Contestant = InputBox("Enter your name.")

frmJeopardy.lblName.Caption = Contestant

frmJeopardy.Show
frmWelcome.Hide

Open App.Path & "\Jeopardy.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Headings(CTR)
    Loop
Close #1

frmJeopardy.lblCategory1.Caption = Headings(1)
frmJeopardy.lblCategory2.Caption = Headings(2)
frmJeopardy.lblCategory3.Caption = Headings(3)
frmJeopardy.lblCategory4.Caption = Headings(4)
frmJeopardy.lblCategory5.Caption = Headings(5)
frmJeopardy.lblCategory6.Caption = Headings(6)
End Sub
