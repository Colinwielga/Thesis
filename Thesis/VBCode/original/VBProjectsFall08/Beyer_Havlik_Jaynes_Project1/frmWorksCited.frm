VERSION 5.00
Begin VB.Form frmWorksCited 
   BackColor       =   &H8000000D&
   Caption         =   "Works Cited"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   BeginProperty Font 
      Name            =   "Candara"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Main Menu"
      Height          =   1215
      Left            =   8640
      TabIndex        =   2
      Top             =   5280
      Width           =   1815
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H80000013&
      Height          =   3855
      Left            =   480
      ScaleHeight     =   3795
      ScaleWidth      =   9915
      TabIndex        =   1
      Top             =   360
      Width           =   9975
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "Show Works Cited"
      Height          =   1575
      Left            =   3960
      TabIndex        =   0
      Top             =   4320
      Width           =   2175
   End
End
Attribute VB_Name = "frmWorksCited"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project name: CSB/SJU Jeopardy
'Form name: frmRules
'Authors: Emma Jaynes, Lindsay Havlik, Brooke Beyer
'Date written: 10/26/08
'Objective: gives credit to the works used to complete this project

Private Sub cmdBack_Click()
'goes back to main menu
frmWorksCited.Hide
frmMainMenu.Show

End Sub

Private Sub cmdView_Click()
'reads file and prints info in picture box
Dim Citations(1 To 100) As String, CTR As Integer, N As Integer

picOutput.Cls

CTR = 0

Open App.Path & "\workscited.txt" For Input As #1

Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Citations(CTR)
Loop
Close #1

picOutput.Print Tab(27); "Works Cited"
picOutput.Print "*****************************************************************************"

For N = 1 To CTR
    picOutput.Print Citations(N)
Next N

End Sub


