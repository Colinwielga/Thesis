VERSION 5.00
Begin VB.Form FrmOptions 
   Caption         =   "Hang-Man Options"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   Picture         =   "FrmOptions.frx":0000
   ScaleHeight     =   3780
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdBackHome 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back to Home Page"
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   3975
   End
   Begin VB.CommandButton CmdSearchBank 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search Word Bank"
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   3975
   End
   Begin VB.CommandButton CmdGetDirections 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Directions For Play"
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   3975
   End
End
Attribute VB_Name = "FrmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Hang Man
'Form Name: FrmOptions
'Authors: Breanna Riskey and Heidi Golliet
'Date Completed: Monday, March 31st
'Objective: The purpose of this form is to allow the user to
'access directions, search word bank for word, or go back to the main form.

Option Explicit

Private Sub CmdGetDirections_Click()
    FrmHome.Visible = False
    FrmOptions.Visible = False
    FrmPlayGame.Visible = False
    FrmHowTo.Visible = True
End Sub


Private Sub CmdBackHome_Click()
    FrmHome.Visible = True
    FrmOptions.Visible = False
    FrmPlayGame.Visible = False
End Sub

Private Sub CmdSearchBank_Click()

Dim WordBank(1 To 100) As String
Dim CTR As Integer
CTR = 0

Open App.Path & "/wordbank.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, WordBank(CTR)
Loop

Close #1

'This is where the user has the option to search the word bank for a word as inputted by them

Dim WordSearch As String
WordSearch = Trim(UCase(InputBox("What word would you like to search for?", , "Word Bank Search")))

Dim Found As Boolean, N As Integer
Found = False
N = 1

Do Until N >= 34
    If WordSearch = Trim(UCase(WordBank(N))) Then
        Found = True
    End If
    N = N + 1
Loop

If Found = True Then
    MsgBox "Yes, we have " & WordSearch & " in the Word Bank.", , "Word Search"
Else
    MsgBox "Sorry, there is no match", , "Word Search"
End If

End Sub
