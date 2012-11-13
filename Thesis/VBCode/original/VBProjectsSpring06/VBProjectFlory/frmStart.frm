VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H00C00000&
   Caption         =   "Starting Attendance"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   ScaleHeight     =   10035
   ScaleWidth      =   11985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H000000FF&
      Caption         =   "All Teams"
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit"
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton cmdIndividual 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Individual Teams"
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H000000FF&
      Caption         =   "Load Attendance Data"
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   5415
      Left            =   2520
      Picture         =   "frmStart.frx":0000
      Top             =   240
      Width           =   8940
   End
   Begin VB.Label lblTylerFlory 
      Alignment       =   2  'Center
      Caption         =   "2005 Major League Baseball Attendance By Tyler Flory"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1800
      TabIndex        =   3
      Top             =   6240
      Width           =   8175
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'to go to the team form
Private Sub cmdAll_Click()
    frmStart.Visible = False
    frmAllTeams.Visible = True
End Sub
'to quit the program
Private Sub cmdExit_Click()
    End
End Sub

'to go to the individual form
Private Sub cmdIndividual_Click()
    frmStart.Visible = False
    frmIndividualTeam.Visible = True
End Sub
'to load the text file
Public Sub cmdLoad_Click()
    Dim pos As Integer
    pos = 0
    
    Open App.Path & "\2005.txt" For Input As #1
    
    Do Until EOF(1)
        pos = pos + 1
        Input #1, Team(pos), HomeGames(pos), HomeTotal(pos), HomeAverage(pos), HomePercent(pos), AwayGames(pos), AwayAverage(pos), AwayPercent(pos), TotalGames(pos), TotalAverage(pos), TotalPercent(pos)
    Loop
    Close #1
    
    Size = pos
End Sub



