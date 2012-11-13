VERSION 5.00
Begin VB.Form frmIndividualTeam 
   BackColor       =   &H0080C0FF&
   Caption         =   "Individual Teams"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Clear"
      Height          =   495
      Left            =   360
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Back"
      Height          =   735
      Left            =   7560
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdRank 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Rank Home Attendance"
      Height          =   495
      Left            =   360
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmdDisplayTotal 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Display Total Stats"
      Height          =   495
      Left            =   360
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton cmdDisplayAway 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Display Away Stats"
      Height          =   495
      Left            =   360
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Exit"
      Height          =   735
      Left            =   9360
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelectTeam 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Select a Team"
      Height          =   495
      Left            =   360
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdDisplayHome 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Display Home Stats"
      Height          =   495
      Left            =   360
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      Height          =   5775
      Left            =   2400
      ScaleHeight     =   5715
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label lblbonds 
      Alignment       =   2  'Center
      Caption         =   "Steriods Taste Like Candy!"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7320
      TabIndex        =   10
      Top             =   5520
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   5280
      Left            =   7080
      Picture         =   "frmShow.frx":0000
      Top             =   120
      Width           =   3810
   End
   Begin VB.Label lblTylerFlory 
      Caption         =   "Tyler Flory"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   8040
      Width           =   1455
   End
End
Attribute VB_Name = "frmIndividualTeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Pick As String
Dim J As Integer
'to go back to the start form
Private Sub cmdBack_Click()
    frmIndividualTeam.Visible = False
    frmStart.Visible = True
End Sub
'clear the picture box
Private Sub cmdClear_Click()
    picResults.Cls
    
End Sub
'to display a teams away stats
Private Sub cmdDisplayAway_Click()
    picResults.Print "AwayGames", "AwayAvg", "AwayPercent"
    picResults.Print AwayGames(J), AwayAverage(J), AwayPercent(J)
End Sub
'to display a teams home stats
Private Sub cmdDisplayHome_Click()
        picResults.Print "HomeGames", "HomeTotal", "HomeAverage"
        picResults.Print HomeGames(J), HomeTotal(J), HomeAverage(J)
End Sub
'to display a teams total stats
Private Sub cmdDisplayTotal_Click()
    picResults.Print "TotalGames", "TotalAverage", "TotalPercent"
    picResults.Print TotalGames(J), TotalAverage(J), TotalPercent(J)
End Sub
'quit the program
Private Sub cmdExit_Click()
    End
End Sub

'to show how good the team is doing on their home attendance
Private Sub cmdRank_Click()
    Select Case HomePercent(J)
    Case Is >= 95
        picResults.Print "Amazing"
    Case Is >= 90
        picResults.Print "Really Good"
    Case Is >= 85
        picResults.Print "Good"
    Case Is >= 80
        picResults.Print "Not Bad"
    Case Is >= 70
        picResults.Print "Mediocre"
    Case Is >= 50
        picResults.Print "Need to Work on it"
    Case Is >= 0
        picResults.Print "Does anybody come?"
    End Select
End Sub
'to be able to pick a team to examine
Private Sub cmdSelectTeam_Click()

Dim found As Boolean
found = False
J = 0
    
    Pick = InputBox("Pick a Team", "Pick")
    
    Do While (found = False And J < Size)
    J = J + 1
    
        If (Pick = Team(J)) Then
            found = True
            picResults.Print "Your Team is "; Team(J)
        End If
    
    Loop
    If found = False Then
MsgBox "Error: Please check team name!", , "Error!!"
    End If
End Sub

