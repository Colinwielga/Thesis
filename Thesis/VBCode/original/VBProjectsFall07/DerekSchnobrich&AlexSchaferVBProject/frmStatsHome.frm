VERSION 5.00
Begin VB.Form frmStatsHome 
   BackColor       =   &H000000FF&
   Caption         =   "Stats"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7665
   FillColor       =   &H000000FF&
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMain 
      Caption         =   "Return to Main Menu"
      Height          =   975
      Left            =   5400
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter Game Statistics"
      Height          =   975
      Left            =   3000
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdSeason 
      Caption         =   "View Cumulative Season Stats"
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblStatsHome 
      BackColor       =   &H000000FF&
      Caption         =   "Statistics"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmStatsHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Takes User to appropriate form
Private Sub cmdEnter_Click()
frmStatsHome.Hide
frmStatsEnter.Show
End Sub
'Returns user to main menu
Private Sub cmdMain_Click()
    frmStatsHome.Hide
    frmHome.Show
End Sub
'Takes User to appropriate form
Private Sub cmdSeason_Click()
frmStatsHome.Hide
frmSeasonStats.Show
End Sub
