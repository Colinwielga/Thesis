VERSION 5.00
Begin VB.Form frmalbuquerqueregion 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4755
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Data"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   1680
      ScaleHeight     =   3195
      ScaleWidth      =   2955
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton frmsort 
      Caption         =   "Sort By Rank"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmddisplay 
      Caption         =   "Display"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Created By: Jeff Amble"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   3360
      Width           =   1695
   End
End
Attribute VB_Name = "frmalbuquerqueregion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form is meant to enable the user to view the teams in the'
'region, as they appear in the file and by rank'
Option Explicit
Dim Names As String
Dim Ranks As Integer
Dim size As Single
Dim TempRank As Integer
Dim TempNames As String
Dim Pass, pos As Integer
'This button enables the user to go back to the main page'
Private Sub cmdback_Click()
    frmalbuquerqueregion.Visible = False
    frmschoolsranksmain.Visible = True
End Sub
'This button enables the user to display the teams and to display'
'the teams when sorted'
Private Sub cmddisplay_Click()
    picresults.Cls
    For pos = 1 To size
        picresults.Print RankArray(pos), NamesArray(pos)
    Next pos
End Sub
'This button enables the user to exit the program'
Private Sub cmdexit_Click()
    End
End Sub
'This button enables the user to load the team information from'
'the corresponding text document'
Private Sub cmdLoad_Click()
    Dim pos As Integer
    Open App.Path & "\albuquerqueregion.txt" For Input As #1
    pos = 0
        Do Until EOF(1)
            pos = pos + 1
            Input #1, NamesArray(pos), RankArray(pos)
        Loop
    Close #1
    size = pos
End Sub
'This button enables the user to sort the teams by rank'
Private Sub frmsort_Click()
    For Pass = 1 To size - 1
        For pos = 1 To size - Pass
            If RankArray(pos) > RankArray(pos + 1) Then
                   TempRank = RankArray(pos)
                   RankArray(pos) = RankArray(pos + 1)
                   RankArray(pos + 1) = TempRank
                   
                   TempNames = NamesArray(pos)
                   NamesArray(pos) = NamesArray(pos + 1)
                   NamesArray(pos + 1) = TempNames
            End If
        Next pos
    Next Pass
End Sub
'This enables the user to view the desired information'
Private Sub picresults_Click()

End Sub
