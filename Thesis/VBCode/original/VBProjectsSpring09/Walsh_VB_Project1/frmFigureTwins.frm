VERSION 5.00
Begin VB.Form frmFigureTwins 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Figure Out the Statistics"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13395
   LinkTopic       =   "Form1"
   Picture         =   "frmFigureTwins.frx":0000
   ScaleHeight     =   8940
   ScaleWidth      =   13395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   3480
      TabIndex        =   10
      Top             =   8040
      Width           =   1815
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Main Menu"
      Height          =   615
      Left            =   840
      TabIndex        =   9
      Top             =   8040
      Width           =   2175
   End
   Begin VB.CommandButton cmdRank 
      Caption         =   "See where they rank against each other!"
      Height          =   735
      Left            =   840
      TabIndex        =   8
      Top             =   7080
      Width           =   2175
   End
   Begin VB.CommandButton cmdViewAll 
      Caption         =   "View all Statistics Together!"
      Height          =   735
      Left            =   840
      TabIndex        =   7
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton cmdOPS 
      Caption         =   "On Base plus Slugging (OPS)"
      Height          =   735
      Left            =   840
      TabIndex        =   6
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton cmdSLG 
      Caption         =   "Slugging Percentage (SLG)"
      Height          =   735
      Left            =   840
      TabIndex        =   5
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton cmdOBP 
      Caption         =   "On Base Percentage (OBP)"
      Height          =   735
      Left            =   840
      TabIndex        =   4
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton cmdBA 
      Caption         =   "Batting Averages (BA)"
      Height          =   735
      Left            =   840
      TabIndex        =   2
      Top             =   2280
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      Height          =   6855
      Left            =   3480
      ScaleHeight     =   6795
      ScaleWidth      =   8715
      TabIndex        =   1
      Top             =   360
      Width           =   8775
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View Stats"
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblFigureOut 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click on a button below to calculate and view each Statistic"
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   1200
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
End
Attribute VB_Name = "frmFigureTwins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim J As Integer

'Baseball Batting Statistics
'frmFigureTwins
'Aaron Walsh
'March 24, 2009
'This program will figure out the various batting statistics like BA, OPS, OBP, and SLG



Private Sub cmdBA_Click()
    'this button finds the batting average of the Twins Players
    picResults.Cls
    picResults.Print "Player", "Batting Average"
    picResults.Print "*************************************"
    For J = 1 To 9
        BA(J) = H(J) / AB(J)
        picResults.Print TwinsNames(J), FormatNumber(BA(J), 3)
    Next J
End Sub

Private Sub cmdBack_Click()
    frmFigureTwins.Hide
    frmInitialform.Show
End Sub

Private Sub cmdOBP_Click()
'this button finds the on base percentage of the Twins Players
    picResults.Cls
    picResults.Print "Player", "On Base Percentage"
    picResults.Print "****************************************"
    For J = 1 To 9
        OBP(J) = (H(J) + BB(J) + HBP(J)) / (AB(J) + BB(J) + HBP(J) + SF(J))
        picResults.Print TwinsNames(J), FormatNumber(OBP(J), 3)
    Next J
End Sub

Private Sub cmdOPS_Click()
'this button finds the OPS of the Twins Players
    picResults.Cls
    picResults.Print "Player", "On Base plus Slugging"
    picResults.Print "********************************************"
    For J = 1 To 9
        OPS(J) = OBP(J) + SLG(J)
        picResults.Print TwinsNames(J), FormatNumber(OPS(J), 3)
    Next J
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdRank_Click()
    frmFigureTwins.Hide
    frmSortTwins.Show
End Sub

Private Sub cmdSLG_Click()
'this button finds the slugging % of the Twins Players
    picResults.Cls
    picResults.Print "Player", "Slugging Percentage"
    picResults.Print "********************************************"
    For J = 1 To 9
        SLG(J) = TB(J) / AB(J)
        picResults.Print TwinsNames(J), FormatNumber(SLG(J), 3)
    Next J
End Sub

Private Sub cmdView_Click()
'this button views the various stats of the Twins Players
    picResults.Cls
    picResults.Print "Player", "At Bats", "Hits", "Home Runs", "Total Bases", "Walks", "Hit-by-Pitch", "Sac Flys"
    picResults.Print "*************************************************************************************************************************************"
    For J = 1 To 9
        picResults.Print TwinsNames(J), AB(J), H(J), HR(J), TB(J), BB(J), HBP(J), SF(J)
    Next J
End Sub

Private Sub cmdViewAll_Click()
    For J = 1 To 9
        SLG(J) = TB(J) / AB(J)
        BA(J) = H(J) / AB(J)
        OBP(J) = (H(J) + BB(J) + HBP(J)) / (AB(J) + BB(J) + HBP(J) + SF(J))
        OPS(J) = OBP(J) + SLG(J)
    Next J
    picResults.Cls
    picResults.Print "Player", "BA", "OBP", "SLG", "OPS"
    picResults.Print "********************************************************************************"
    For J = 1 To 9
        picResults.Print TwinsNames(J), FormatNumber(BA(J), 3), FormatNumber(OBP(J), 3), FormatNumber(SLG(J), 3), FormatNumber(OPS(J), 3)
    Next J
End Sub


