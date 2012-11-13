VERSION 5.00
Begin VB.Form frmIntro 
   Caption         =   "The Art of Procrastination"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   Picture         =   "FirstProject.frx":0000
   ScaleHeight     =   5250
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetStarted 
      Caption         =   "Get Started Now!"
      Height          =   975
      Left            =   1560
      TabIndex        =   0
      Top             =   3480
      Width           =   3255
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetStarted_Click()

'this uploads the parameters for each interveral of
'of homework a user could have, then messages them
'whether or not they should be procrastinating
'it also ask the user for their name and the number of
'hours of homework they have and takes them to the main
'screen

Dim CTR As Integer
Dim Pos As Integer
Dim Found As Boolean
frmIntro.Hide
Found = False
Open App.Path & "\Procrastination.txt" For Input As #1
CTR = 0
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, Hours(CTR), Responce(CTR)
    Loop
Name1 = InputBox("Enter your name")
Homework = InputBox("How many hours of homework do you have?")
For Pos = 1 To CTR
    If Found = False Then
        If Hours(Pos) > Homework Then
            MsgBox (Responce(Pos))
            Found = True
        End If
    End If
Next Pos
If Found = False Then
    MsgBox (Responce(5))
End If
frmMainScreen.Visible = True
End Sub
