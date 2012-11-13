VERSION 5.00
Begin VB.Form frmScores 
   BackColor       =   &H00FF0000&
   Caption         =   "Scores"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9120
      TabIndex        =   18
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   17
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "Calculate Fantasy Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      TabIndex        =   16
      Top             =   5040
      Width           =   1815
   End
   Begin VB.PictureBox pbxResults 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3120
      ScaleHeight     =   1755
      ScaleWidth      =   5595
      TabIndex        =   15
      Top             =   5040
      Width           =   5655
   End
   Begin VB.TextBox txtTwo 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8400
      TabIndex        =   13
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox txtInt 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox txtRushyd 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox txtRecyd 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox txtTd 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8400
      TabIndex        =   9
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox txtPasstd 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8400
      TabIndex        =   8
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtpassyd 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lbltwo 
      BackColor       =   &H00FF0000&
      Caption         =   "Two Point Conversion (Pass, Run, or Catch)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5880
      TabIndex        =   14
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label lblInt 
      BackColor       =   &H00FF0000&
      Caption         =   "Passing Interceptions:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblRush 
      BackColor       =   &H00FF0000&
      Caption         =   "Rushing Yards:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label lblRec 
      BackColor       =   &H00FF0000&
      Caption         =   "Recieving Yards:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label lblTD 
      BackColor       =   &H00FF0000&
      Caption         =   "Rushing or Recieving TD's:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5880
      TabIndex        =   3
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label lblPasstd 
      BackColor       =   &H00FF0000&
      Caption         =   "Passing TD's:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblPass 
      BackColor       =   &H00FF0000&
      Caption         =   "Passing Yards:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Enter Stats For Any Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmScores (frmScores.frm)
'Purpose of this form is to allow the user to input statistics
'of any player, and thier fantasy score will eb displayed

'Scoring System:
'1 pt. for every 25 yds.passing, 10 yds. recieving, or 10 yds. Rushing
'6 pts for every Rush or Rec. Touchdown, 4 pts for Passing Touchdown
'-1 pt. for every Interception thrown
'2 pts. for any 2 pt. conversion


Option Explicit



Private Sub cmdClear_Click()
    pbxResults.Cls
'clears the picturebox

End Sub

Private Sub cmdCompute_Click()
Dim PassYards, PassInt, RushYards, RecYards, PassTD, TD, TwoPoint, Points As Integer
'declares all the variables as Integers


    PassYards = txtpassyd.Text
    PassInt = txtInt.Text
    RushYards = txtRushyd.Text
    RecYards = txtRecyd.Text
    PassTD = txtPasstd.Text
    TD = txtTd.Text
    TwoPoint = txtTwo.Text
'each variable is the value of its textbox


Points = (Int(PassYards / 25)) + (PassInt * -2) + (Int(RushYards / 10)) + (Int(RecYards / 10)) + (PassTD * 4) + (TD * 6) + (TwoPoint * 2)
'Int( allows the remaining value to be truncated

pbxResults.Print "Your Player's Fantasy Score"
pbxResults.Print "***************************"
pbxResults.Print Points

If Points > 30 Then
    pbxResults.Print "Awesome Score!"
ElseIf Points > 20 Then
    pbxResults.Print "Very Good Score"
ElseIf Points > 10 Then
    pbxResults.Print "Okay Score"
ElseIf Points > 0 Then
    pbxResults.Print "Not that good"
ElseIf Points = 0 Then
    pbxResults.Print "No Score"
ElseIf Points < 0 Then
    pbxResults.Print "Horrible Score, YOU LOST POINTS!!"
End If

'prints title and pt. value in the picture box, also tells if score is good, bad, etc.

End Sub

Private Sub cmdQuit_Click()
'ends program

    End
End Sub

