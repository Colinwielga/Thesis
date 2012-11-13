VERSION 5.00
Begin VB.Form frmStandings 
   Caption         =   "Current NBA Standing"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdView 
      BackColor       =   &H00FF0000&
      Caption         =   "View Current NBA Standings"
      Height          =   1455
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Go Back"
      Height          =   1575
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   2775
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9975
      Left            =   -960
      ScaleHeight     =   9915
      ScaleWidth      =   13635
      TabIndex        =   1
      Top             =   0
      Width           =   13695
   End
   Begin VB.Label lblName 
      Caption         =   "By: Chad Henfling"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12600
      TabIndex        =   3
      Top             =   4800
      Width           =   2415
   End
End
Attribute VB_Name = "frmStandings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Minnesota Timberwolves Center (MinnesotaTimberwovlesbyChadHenfling.vbp)
'Main Form (frmStadings.frm)
'Chad Henfling
'Created March 23, 2006
'This form shows the current NBA standings in division and throughout the league.
Option Explicit
Dim Standings(1 To 100) As String
Dim pos As Integer


Private Sub cmdBack_Click()
    'go back to main form
    frmStandings.Visible = False
    frm1.Visible = True
End Sub

Private Sub cmdView_Click()
    'Opening file and reading information
    Open App.Path & "\Standings.txt" For Input As #3
    pos = 0
    'reading entire file and printing the entire string
    Do Until EOF(3)
        pos = pos + 1
        Input #3, Standings(pos)
        picOutput.Print Standings(pos)
    Loop
    Close #3
End Sub

