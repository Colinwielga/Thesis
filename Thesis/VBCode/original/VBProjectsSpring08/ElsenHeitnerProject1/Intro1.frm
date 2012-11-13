VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H000000FF&
   Caption         =   "Intro Form"
   ClientHeight    =   6210
   ClientLeft      =   3945
   ClientTop       =   3135
   ClientWidth     =   8550
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   Picture         =   "Intro1.frx":0000
   ScaleHeight     =   6210
   ScaleWidth      =   8550
   Begin VB.PictureBox picResults 
      Height          =   3135
      Left            =   600
      Picture         =   "Intro1.frx":3186
      ScaleHeight     =   3075
      ScaleWidth      =   7515
      TabIndex        =   6
      Top             =   960
      Width           =   7575
   End
   Begin VB.CommandButton Submit 
      Caption         =   "Submit"
      Height          =   615
      Left            =   7080
      TabIndex        =   5
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   615
      Left            =   5280
      TabIndex        =   4
      Text            =   "Enter your name here"
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "Help Me Choose"
      Height          =   615
      Left            =   4560
      TabIndex        =   2
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Cantidate Biographies"
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00000000&
      Caption         =   "Lose (Quit)"
      Height          =   615
      Left            =   6720
      MaskColor       =   &H00000000&
      TabIndex        =   0
      Top             =   5400
      Width           =   1695
   End
   Begin VB.OLE OLE1 
      Class           =   "Package"
      Height          =   735
      Left            =   240
      OleObjectBlob   =   "Intro1.frx":15156
      TabIndex        =   3
      Top             =   4320
      Width           =   1935
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PROJECT: Choose or Lose: Election Perfection
'FORM: Intro Form(frmIntro.frm)
'AUTHOR:  Nick Elsen and Andrew Heitner
'DATE:  March 12, 2008
'PURPOSE:  The overall purpose of this project is to give the user background on the cantidates of the upcoming election and help them choose one with they most similar views.
'          This form is an intro form that promts the user for their name and then allows the user to go to the next form or quit.


Option Explicit

'Takes you to the next form to start the choosing process
Private Sub cmdChoose_Click()
If Module1.UsersName = "notset" Then
    MsgBox "Please enter your name, then press Submit"
Else
frmChoose.Show
frmIntro.Hide
    MsgBox "The four cantidates have all taken their stances. " & Chr(13) & "Here are five topics that they have made a decision on. " & Chr(13) & "Click each one and choose which response you agree with most! " & Chr(13) & "When you have answered all five questions, " & Chr(13) & "click on the button to see who your views agree with the most!", , "Views"
End If

End Sub

'Takes you to the cantidates form
Private Sub cmdNext_Click()
Select Case Module1.UsersName
    Case Is = "notset"
        MsgBox "Please enter your name, then press Submit"
    Case Is <> "notset"
        frmIntro.Hide
        frmCantidates.Show

End Select

End Sub

'Ends the program
Private Sub cmdQuit_Click()
End
End Sub

'Sets the Users Input name to the public variable
Private Sub Submit_Click()
Module1.UsersName = frmIntro.txtName
End Sub

