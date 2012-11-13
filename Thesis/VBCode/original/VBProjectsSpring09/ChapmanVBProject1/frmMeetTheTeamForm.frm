VERSION 5.00
Begin VB.Form MeetTheTeam 
   Caption         =   "Form1"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdToalGE 
      Caption         =   "Total Goals and Ejections"
      Height          =   1095
      Left            =   5040
      TabIndex        =   4
      Top             =   4560
      Width           =   2655
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort by name"
      Height          =   1215
      Left            =   5040
      TabIndex        =   3
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search for someone"
      Height          =   1215
      Left            =   5040
      TabIndex        =   2
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read Arrays"
      Height          =   1095
      Left            =   5040
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.PictureBox picResults 
      Height          =   6495
      Left            =   480
      ScaleHeight     =   6435
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "MeetTheTeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Names(1 To 15) As String, Goals(1 To 15) As Integer
Dim Ejections(1 To 15) As Integer, Ctr As Integer

Private Sub cmdRead_Click()

Ctr = 0

Open App.Path & "\TheTeam.txt" For Input As #1

picResults.Print "Name", Tab(20); "Goals", Tab(30); "Ejections"

Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Names(Ctr), Goals(Ctr), Ejections(Ctr)
    picResults.Print Names(Ctr), Goals(Ctr), Ejections(Ctr)
Loop

Close #1

End Sub
