VERSION 5.00
Begin VB.Form FrmElection1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5760
      Width           =   2415
   End
   Begin VB.CommandButton CmdResults 
      BackColor       =   &H00FFFFFF&
      Caption         =   "The Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   2655
   End
   Begin VB.CommandButton CmdCandidates 
      BackColor       =   &H00FF0000&
      Caption         =   "The Candidates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   2535
   End
   Begin VB.PictureBox PicButton 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   1200
      Picture         =   "FrmElection1.frx":0000
      ScaleHeight     =   4455
      ScaleWidth      =   6015
      TabIndex        =   1
      Top             =   720
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "The 2008 Presidential Election"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "FrmElection1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'Election Project
    'FrmElection1
    'Ian Bouman
    'Written on 3/10
    'The purpose of this form is to direct the user in two different
    'directions - either to the results or to the candidates.
    'The overall purpose of this project is to show the user the
    'results of the 2008 presidential election to this date and to
    'allow the user to see the results - both delegates and the
    'percentage of popular vote per candidate - in each state.
Private Sub CmdCandidates_Click()
FrmElection1.Hide
FrmElection2.Show
End Sub

Private Sub CmdQuit_Click()
End
End Sub

Private Sub CmdResults_Click()
FrmElection1.Hide
FrmElection3.Show
End Sub
