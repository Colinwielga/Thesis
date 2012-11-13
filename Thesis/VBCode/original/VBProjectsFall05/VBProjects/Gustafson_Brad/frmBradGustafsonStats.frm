VERSION 5.00
Begin VB.Form frmBradGustafsonStats 
   Caption         =   "BradGustafsonStats"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   Picture         =   "frmBradGustafsonStats.frx":0000
   ScaleHeight     =   6345
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton cmdInterceptionsStats 
      Caption         =   "Interceptions"
      Height          =   735
      Left            =   6360
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdcmdIndDefenseStats 
      Caption         =   "Individual Stats"
      Height          =   735
      Left            =   6360
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdPassingStats 
      Caption         =   "Passing Stats"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdReceivingStats 
      Caption         =   "Receiving Stats"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdRushingStats 
      Caption         =   "Rushing Stats"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdTeamStats 
      Caption         =   "Team Stats"
      Height          =   735
      Left            =   3240
      TabIndex        =   0
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label lblDefenseStats 
      BackColor       =   &H00FF0000&
      Caption         =   "DEFENSE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblOffensiveStats 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "OFFENSE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmBradGustafsonStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click() 'This button brings you back to the main page of my project'
    frmBradGustafsonStats.Hide
End Sub

Private Sub cmdIndDefenseStats_Click() 'This button shows a form containing the Individual Defense stats for the Cowboys'
    frmIndividualStats.Show
End Sub

Private Sub cmdInterceptionsStats_Click() 'This button shows a form containing the Interception stats for the Cowboys'
    frmInterceptionStats.Show
End Sub

Private Sub cmdPassingStats_Click() 'This button shows a form containing the passing stats for the Cowboys'
    frmPassingStats.Show
   
End Sub

Private Sub cmdReceivingStats_Click() 'This button shows a form containing the receiving stats for the Cowboys'
    frmReceivingStats.Show
End Sub

Private Sub cmdRushingStats_Click() 'This button shows the form containing the rushing stats for the Cowboys'
    frmRushingStats.Show
End Sub

Private Sub cmdTeamStats_Click() 'This button shows the form containing the Team stats for the Cowboys'
    frmTeamStats.Show
End Sub

