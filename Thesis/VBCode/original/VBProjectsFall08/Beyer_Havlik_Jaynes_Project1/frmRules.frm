VERSION 5.00
Begin VB.Form frmRules 
   BackColor       =   &H8000000D&
   Caption         =   "Rules of the Game"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOutput 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   2280
      ScaleHeight     =   5955
      ScaleWidth      =   9195
      TabIndex        =   4
      Top             =   600
      Width           =   9255
   End
   Begin VB.CommandButton cmdMainMenu 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   3
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdFinal 
      Caption         =   "Final Jeopardy Rules"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdDoubles 
      Caption         =   "Daily Double Rules"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Display General Gameplay Rules"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      MaskColor       =   &H00808080&
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project name: CSB/SJU Jeopardy
'Form name: frmRules
'Authors: Emma Jaynes, Lindsay Havlik, Brooke Beyer
'Date written: 10/26/08
'Objective: form displays various rules of the game when specific buttons are clicked
'Comments: 1.General Rules  2.Double Jeopardy Rules  3.Final Jeopardy Rules  4.Back to Main Menu

Dim CTR As Integer, N As Integer

Private Sub cmdDisplay_Click()
'reads file and displays info in picture box
Dim Rules(1 To 100) As String

picOutput.Cls

Open App.Path & "\generalrules.txt" For Input As #1

CTR = 0

Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Rules(CTR)
Loop
Close #1

picOutput.Print Tab(18); "General Rules of the Game"
picOutput.Print "*******************************************************************"

For N = 1 To CTR
    picOutput.Print Rules(N)
Next N

End Sub

Private Sub cmdDoubles_Click()
'reads file and displays info in picture box
Dim DoubleRules(1 To 100) As String

picOutput.Cls

Open App.Path & "\doublerules.txt" For Input As #1

CTR = 0

Do Until EOF(1)
    CTR = CTR + 1
    Input #1, DoubleRules(CTR)
Loop
Close #1

picOutput.Print Tab(18); "Double Jeopardy Rules"
picOutput.Print "*******************************************************************"

For N = 1 To CTR
    picOutput.Print DoubleRules(N)
Next N
End Sub


Private Sub cmdFinal_Click()
'reads file and displays info in picture box
Dim FinalRules(1 To 100) As String

picOutput.Cls

Open App.Path & "\finalrules.txt" For Input As #1

CTR = 0

Do Until EOF(1)
    CTR = CTR + 1
    Input #1, FinalRules(CTR)
Loop
Close #1

picOutput.Print Tab(18); "Final Jeopardy Rules"
picOutput.Print "*******************************************************************"

For N = 1 To CTR
    picOutput.Print FinalRules(N)
Next N
End Sub

Private Sub cmdMainMenu_Click()
'goes back to main menu
frmRules.Hide
frmMainMenu.Show

End Sub


