VERSION 5.00
Begin VB.Form frmhw1 
   Caption         =   "Form1"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   1320
      TabIndex        =   2
      Top             =   3600
      Width           =   2535
   End
   Begin VB.PictureBox picResults 
      Height          =   1215
      Left            =   600
      ScaleHeight     =   1155
      ScaleWidth      =   4875
      TabIndex        =   1
      Top             =   1680
      Width           =   4935
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "click to enter information"
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmhw1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim runningTotal As Integer

Private Sub CmdInfo_Click()
Dim Major As String, SportsTeam As String, Classes As Integer, GPA As Single

'Get a Major from the user using an InputBox
Major = InputBox("Please enter your major")

If Major = "Management" Then
MsgBox ("That's my major too!")
runningTotal = 1

End If

'Get a favorite sports team from the user using an InputBox
SportsTeam = InputBox("Please enter your favorite sports team")

If SportsTeam = "Celtics" Then
MsgBox ("That's my favorite sports team too!")
runningTotal = runningTotal + 1

End If

'Get number of classes from the user using an InputBox
Classes = InputBox("Please enter the number of classes you are curently enrolled in")

If Classes = "5" Then
MsgBox ("I'm taking 5 classes too!")
runningTotal = runningTotal + 1
End If

'Get a desired GPA from the user usin an InputBox
GPA = InputBox("Please enter your desired GPA for this term")

If GPA = "4.0" Then
MsgBox ("I hope I get a 4.0 GPA too!")
runningTotal = 1
ElseIf GPA > "4.0" Then
MsgBox ("That does not make sense!")
ElseIf GPA < 0 Then
MsgBox ("That does not make sense!")
End If

If runningTotal = 0 Then
picResults.Print "We have nothing in common in regards to these questions."
ElseIf runningTotal = 1 Then
picResults.Print "Well, at least we have one thing in common."
ElseIf runningTotal = 2 Then
picResults.Print "Cool, 2 matches!"
ElseIf runningTotal = 3 Then
picResults.Print "Wow!"
ElseIf runningTotal = 4 Then
picResults.Print "We are on the same wavelength! Sweet!"
End If

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub Picture1_Click()

End Sub

