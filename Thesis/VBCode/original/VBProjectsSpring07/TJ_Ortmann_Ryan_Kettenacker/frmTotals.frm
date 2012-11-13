VERSION 5.00
Begin VB.Form frmTotals 
   BackColor       =   &H000080FF&
   Caption         =   "Are You Winning!!!"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoToRankings 
      Caption         =   "Go To Rankings"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      TabIndex        =   5
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton cmdGoBackFinals 
      Caption         =   "Go Back To Finals"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      TabIndex        =   3
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      TabIndex        =   2
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdloadname 
      Caption         =   "CLICK HERE TO LOAD NAME AND SCORES "
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5280
      TabIndex        =   1
      Top             =   2520
      Width           =   2655
   End
   Begin VB.PictureBox picname 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   360
      ScaleHeight     =   4755
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   2040
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   2220
      Left            =   5400
      Picture         =   "frmTotals.frx":0000
      Top             =   120
      Width           =   2250
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Totals From Regions And Final Four"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmTotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' this form is for the user to see there totals from each round and there grand total

Private Sub cmdGoBackFinals_Click()
    frmTotals.Hide                      'allow user to go from Totals form to Finals form
    frmFinals.Show
End Sub

Private Sub cmdGoToRankings_Click()
    frmTotals.Hide                      'allow user to go from Totals form to Rankings form
    frmRankings.Show
End Sub

'Code for printing all results in to picname
Private Sub cmdloadname_Click()
    
    OverallScore = EastTotal + SouthTotal + WestTotal + MidwestTotal + Final4Sum + ChampionshipSum + ChampsSum  'giving the total score a variable
    picname.Cls                                                                                                 'clear picture box before printing
    picname.Print User                                                                                          'print user name
    picname.Print "__________________________________________________"
    picname.Print "Total From Round 1"
    picname.Print EastR1Sum + WestR1Sum + SouthR1Sum + MidwestR1Sum                                             'round 1 total score
    picname.Print "Total From Round 2"
    picname.Print EastR2Sum + WestR2Sum + SouthR2Sum + MidwestR2Sum                                             'round 2 total score
    picname.Print "Total From Round 3"
    picname.Print EastR3Sum + WestR3Sum + SouthR3Sum + MidwestR3Sum                                             'round 3 total score
    picname.Print "Final Four Total"
    picname.Print Final4Sum                                                                                     'final four total
    picname.Print "Championship Total"
    picname.Print ChampionshipSum                                                                               'championship total
    picname.Print "Champion Total"
    picname.Print ChampsSum                                                                                     'Champion sum
    picname.Print "Total Bracket Score"
    picname.Print OverallScore                                                                                  'Total score from each round, final score
    
End Sub

Private Sub cmdQuit_Click()
    End                                         'End Program
End Sub
