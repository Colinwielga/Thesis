VERSION 5.00
Begin VB.Form frmPlayersearch 
   BackColor       =   &H80000012&
   Caption         =   "Form2"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8055
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   5670
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFindplayer 
      Caption         =   "Find an SJU Player"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdBackmenu1 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   3
      Top             =   4920
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00008000&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      FontTransparent =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1875
      ScaleWidth      =   4755
      TabIndex        =   2
      Top             =   720
      Width           =   4815
   End
   Begin VB.TextBox txtPlayername 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   2760
      Picture         =   "Playersearch.frx":0000
      Top             =   3120
      Width           =   1425
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Dan Ruehl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Pat Boerner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Eric Holmgren"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Ted Lauer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Curtis Horton"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Adam Putschoegl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Trevor Beach"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Steve Tacl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ROSTER 2006"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lblEntername 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "*Enter Players Full Name (First Last)"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   5670
      Left            =   4320
      Picture         =   "Playersearch.frx":09FC
      Top             =   0
      Width           =   3750
   End
End
Attribute VB_Name = "frmPlayersearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'2006 MIAC Tennis Tournament Distribution
'Player Search Form
'Blake Heymans
'10/26/06
'Player Search Form Objective
    'This form is intended to show the user the individual match data for the player
    'of the user's choice. First it promts the user to enter a player's name.
    'The user's player is then searched for in two files containing data on
    'matches played in order to see the results of the player in the tournament.
'Pictures were taken from Google image search as well as the Saint John's Univerity web site.

Private Sub cmdBackmenu1_Click()
    'Brings the user back to Title Form
    frmTitle.Show
    frmPlayersearch.Hide
End Sub
Private Sub cmdFindplayer_Click()
Dim Pname As String, P As Integer
Dim Ctr As Integer

    picResults.Cls
    
    'Singles Players Search
    Ctr = 0
    
    Open App.Path & "\Singles.txt" For Input As #1
    
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, Playernames(Ctr), Scores(Ctr), Opponent(Ctr), Winloss(Ctr)
    Loop
    
    Close #1

    Pname = txtPlayername.Text
    'Prints the titles that will be used
    picResults.Print "Scores"; Tab(15); "Opponent"; Tab(30); "Win Or A Loss"; Tab(47); "Doubles Partner"
    picResults.Print "==================================================================================================="
    'Prints the info that the user is searching for from the file 1
    For P = 1 To Ctr
        If Pname = Playernames(P) Then
            picResults.Print Scores(P); Tab(15); Opponent(P); Tab(30); Winloss(P); Tab(47); "Singles Match"
        End If
    Next P
    
    'Now The Doubles Player Search
    Ctr = 0
    
    Open App.Path & "\Doubles.txt" For Input As #2
    
    Do While Not EOF(2)
        Ctr = Ctr + 1
        Input #2, Playernamesd(Ctr), Scoresd(Ctr), Opponentd(Ctr), Winlossd(Ctr), Dpartner(Ctr)
    Loop
    
    Close #2
    'Prints the info that the user is searching for from the file 2
    For P = 1 To Ctr
        If Pname = Playernamesd(P) Then
            picResults.Print Scoresd(P); Tab(15); Opponentd(P); Tab(30); Winlossd(P); Tab(47); Dpartner(P)
        End If
    Next P
End Sub
