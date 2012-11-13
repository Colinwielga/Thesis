VERSION 5.00
Begin VB.Form frminfo 
   BackColor       =   &H00404000&
   Caption         =   "Garrett Sohn"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pichistory 
      Height          =   2295
      Left            =   6360
      Picture         =   "frminfo.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   1755
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
   Begin VB.PictureBox picbracket 
      Height          =   1695
      Left            =   6120
      Picture         =   "frminfo.frx":CC2A
      ScaleHeight     =   1635
      ScaleWidth      =   2235
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton cmdmain 
      Caption         =   "Main Page"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.PictureBox picncaalogo 
      Height          =   2415
      Left            =   3000
      Picture         =   "frminfo.frx":18978
      ScaleHeight     =   2355
      ScaleWidth      =   2475
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label lblclick 
      Caption         =   "Click the picture for information on the tournament"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   480
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
End
Attribute VB_Name = "frminfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'March madness (madness.vbp)
'information form (info.frm)
'Garrett Sohn
'March 24, 2006
'This form gives information by clicking on the picture boxes for different types of information.
Option Explicit
Private Sub cmdmain_Click()
    frminfo.Hide
    frmmadness.Show
End Sub

Private Sub picbracket_Click()
    MsgBox "The bracket is the most vital part of the tournament for the national audience.  This is where people make their office pool wagers by trying to predict who will win each game and become the overall champion.  CBS has made this bracket a household item every spring, while their analysts try and predict who will win. Each team is given a seeding before the tournament begins by a committee.  The committee decides each seeding by the number of games they've won, against who, and their RPI, which is a formula that demonstrates a team's strength of schedule.  The bracket has almost gone out of date because of the increase of internet users creating online pools.", , "The Bracket"
End Sub

Private Sub pichistory_Click()
    MsgBox "The tournament began in 1939 with the format of the tournament changing constantly, accomodating to the amount of teams and mass interest of the tournament.  In 1939, there were only eight teams compared to the 65 teams there are in the tournament today.  The newest edition to the tournament is the play-in game, which has the two lowest seeds play eachother to get into the 64 team tournament. "
End Sub

Private Sub picncaalogo_Click()
    MsgBox "The NCAA Division 1 Men's Basketball Tournament is compiled of the top 65 college teams in the United States. Each team makes the tournament by either winning their conference tournament or recieving an at large bid.  It is a 20-day event throughout the months of March and beginning of April. The tournament is more commonly known by March Madness and the Big Dance.  It has become one of the most prominent sporting events within the United States each year.", , "Tournament Description"
End Sub


