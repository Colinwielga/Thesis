VERSION 5.00
Begin VB.Form frmCareerStats 
   BackColor       =   &H80000008&
   Caption         =   "How Good Was Pete Rose?"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   1335
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdComparePete 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   7080
      Picture         =   "frmCareerStats.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   3015
   End
   Begin VB.CommandButton cmdCompare 
      Enabled         =   0   'False
      Height          =   2055
      Left            =   3600
      Picture         =   "frmCareerStats.frx":AEDE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   3015
   End
   Begin VB.CommandButton cmdReturnMenu 
      BackColor       =   &H000000FF&
      DisabledPicture =   "frmCareerStats.frx":15AA6
      Height          =   1215
      Left            =   10320
      Picture         =   "frmCareerStats.frx":1D882
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton cmdGetStats 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      Picture         =   "frmCareerStats.frx":251C7
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   3015
   End
   Begin VB.PictureBox picResultsCareerStats 
      BackColor       =   &H00C0C0FF&
      Height          =   6015
      Left            =   120
      ScaleHeight     =   5955
      ScaleWidth      =   10155
      TabIndex        =   0
      Top             =   120
      Width           =   10215
   End
   Begin VB.Image picRoseKneel 
      Height          =   8280
      Left            =   10320
      Picture         =   "frmCareerStats.frx":2F54E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4875
   End
End
Attribute VB_Name = "frmCareerStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCompare_Click()
    'shows form HOFstats (Hall of Fame stats) to compare Pete Rose with recent Hall of Famers
    cmdGetStats.Enabled = True
    frmHoFstats.Show
End Sub

Private Sub cmdComparePete_Click()
    'shows form RoseCompared to compare Rose 1-on-1 with the given Hall of Famers
    cmdGetStats.Enabled = True
    frmRoseCompared.Show
End Sub

Private Sub cmdGetStats_Click()
    'displays text from \statsCareer.txt and displays Rose's career stats; prints results
    Dim YR, GP, AB, R, H, DBL, TRI, HR, RBI, BB, SO, SB, CS, BA As Double
    picResultsCareerStats.Cls
    picResultsCareerStats.Print Tab(55); "Pete Rose Career Statistics"; Tab(1);
    picResultsCareerStats.Print " YEAR"; Tab(11); "G"; Tab(21); "AB"; Tab(31); "R"; Tab(41); "H"; Tab(51); "2B"; Tab(60); "3B"; Tab(70); "HR"; Tab(80); "RBI"; Tab(90); "BB"; Tab(100); "SO"; Tab(110); "SB"; Tab(121); "CS"; Tab(131); "BA"
    picResultsCareerStats.Print "***************************************************************************************************************************************************************************";
    Open App.Path & "\statsCareer.txt" For Input As #1
    Do Until EOF(1)
        Input #1, YR, GP, AB, R, H, DBL, TRI, HR, RBI, BB, SO, SB, CS, BA
        picResultsCareerStats.Print
        picResultsCareerStats.Print YR; Tab(10); GP; Tab(20); AB; Tab(30); R; Tab(40); H; Tab(50); DBL; Tab(60); TRI; Tab(70); HR; Tab(80); RBI; Tab(90); BB; Tab(100); SO; Tab(110); SB; Tab(120); CS; Tab(130); Right(BA, 4);
    Loop
    Close #1
    picResultsCareerStats.Print Chr(10); "***************************************************************************************************************************************************************************";
    picResultsCareerStats.Print Chr(10); " 24-Year"; " ****************************************************************************************************************************************************************"; Chr(10);
    picResultsCareerStats.Print " Totals"; Tab(10); "3562"; Tab(20); "14053"; Tab(30); "2165"; Tab(40); "4256"; Tab(50); "746"; Tab(60); "135"; Tab(70); "160"; Tab(80); "1314"; Tab(90); "1566"; Tab(100); "1143"; Tab(110); "198"; Tab(120); "149"; Tab(130); ".303";
    cmdCompare.Enabled = True
    cmdComparePete.Enabled = True
    cmdGetStats.Enabled = False
End Sub

Private Sub cmdReturnMenu_Click()
    'shows menu page and removes this page from visibility
    cmdGetStats.Enabled = True
    frmCareerStats.Hide
    frmMenuPage.Show
End Sub
