VERSION 5.00
Begin VB.Form frmTwins 
   Caption         =   "Opening"
   ClientHeight    =   6900
   ClientLeft      =   2055
   ClientTop       =   3090
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   Picture         =   "OpeningForm.frx":0000
   ScaleHeight     =   6900
   ScaleWidth      =   9930
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit Program"
      Height          =   855
      Left            =   8280
      TabIndex        =   3
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton cmdSeries 
      Caption         =   "World Series Stats"
      Height          =   855
      Left            =   8280
      TabIndex        =   2
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdPitching 
      Caption         =   "Pitching Stats"
      Height          =   855
      Left            =   8280
      TabIndex        =   1
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdHitting 
      Caption         =   "Hitting Stats"
      Height          =   855
      Left            =   8280
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmTwins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: 1987 World Series
'Form name: frmTwins
'Authors: Hans Paul and Cole Wuollet
'Date Written: Tuesday October 31, 2006
'Objective: The overall Objective of this program is to inform the user about the
            '1987 World Series Champion Minnesota Twins.  This program will provide
            'information on Hitting statistics and Pitching Statistics for the team
            'in both the regular season and in the World Series.  This program assumes
            'that the user has basic knowledge about baseball terminology and statistics.
Option Explicit

Private Sub cmdHitting_Click()
    frmTwins.Hide
    frmHitting.Show
    
End Sub

Private Sub cmdpitching_Click()
    frmTwins.Hide
    frmPitching.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSeries_Click()
    frmTwins.Hide
    frmSeries.Show
End Sub

