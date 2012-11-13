VERSION 5.00
Begin VB.Form frmTeam_Info 
   BackColor       =   &H80000013&
   Caption         =   "Form1"
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14055
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   14055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBacktoMain 
      Caption         =   "Back to front page"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6120
      TabIndex        =   6
      Top             =   7680
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   6855
      Left            =   8280
      Picture         =   "Team_Info.frx":0000
      ScaleHeight     =   6795
      ScaleWidth      =   5595
      TabIndex        =   5
      Top             =   1800
      Width           =   5655
   End
   Begin VB.CommandButton cmdRecords 
      Caption         =   "View All-Time Best Performances"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6120
      TabIndex        =   4
      Top             =   4560
      Width           =   2055
   End
   Begin VB.PictureBox picRecords 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   240
      ScaleHeight     =   3555
      ScaleWidth      =   5715
      TabIndex        =   3
      Top             =   5520
      Width           =   5775
   End
   Begin VB.CommandButton cmdTeamInfo 
      Caption         =   "View Team Information"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6120
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
   End
   Begin VB.PictureBox picQuickFacts 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   240
      ScaleHeight     =   3675
      ScaleWidth      =   5715
      TabIndex        =   1
      Top             =   1560
      Width           =   5775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Team Information"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   14055
   End
End
Attribute VB_Name = "frmTeam_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form allows the user to view facts about the history of St. John's Cross Country,
'including the all-time best performances, and other fun facts about St. John's

'takes the user back to the front page
Private Sub cmdBacktoMain_Click()
    frmTeam_Info.Hide
    frmSJU_CC.Show
End Sub

'pressing this button shows the user some of the top performances in the history of St. John's Cross Country
Private Sub cmdRecords_Click()
    picRecords.Print "Name"; Tab(20); "Time"
    picRecords.Print "--------------------------------------------------------"
    picRecords.Print "Brian Smith"; Tab(20); "23:56"
    picRecords.Print "Chris Erichsen"; Tab(20); "24:10"
    picRecords.Print "John Kruger"; Tab(20); "24:14"
    picRecords.Print "John Gathje"; Tab(20); "24:16"
    picRecords.Print "Kelly Fermoyle"; Tab(20); "24:24"
    picRecords.Print "John Cragg"; Tab(20); "24:30"
    picRecords.Print "Chet Boom"; Tab(20); "24:31"
    picRecords.Print "Chuck Ceronsky"; Tab(20); "25:01"
    picRecords.Print "Charlie Mahler"; Tab(20); "25:02"
    picRecords.Print "Joe Metzger"; Tab(20); "25:03"
    
    cmdRecords.Enabled = False
    cmdBacktoMain.Enabled = True
    
End Sub

'pressing this button allows the user to view some basic facts about St. John's and the cross country team
Private Sub cmdTeamInfo_Click()
    picQuickFacts.Print Tab(15); "Saint John's University"
    picQuickFacts.Print Tab(2); "--------------------------------------------------------"
    picQuickFacts.Print "Location: Collegeville, MN"
    picQuickFacts.Print "Enrollment: 1,917"
    picQuickFacts.Print "School Colors: Cardinal and Blue"
    picQuickFacts.Print "Nickname: Johnnies"
    picQuickFacts.Print "Coach: Tim Miles"
    picQuickFacts.Print Tab(2); "--------------------------------------------------------"
    picQuickFacts.Print "MIAC Championships: 15"
    picQuickFacts.Print "NCAA D3 Central Region Championships: 5"
    picQuickFacts.Print "2007 MIAC Champions"
    
    cmdTeamInfo.Enabled = False
    cmdRecords.Enabled = True
    
End Sub

