VERSION 5.00
Begin VB.Form frmMinnesotaWildHockeyTeam 
   BackColor       =   &H00008000&
   Caption         =   "Title"
   ClientHeight    =   7380
   ClientLeft      =   855
   ClientTop       =   1110
   ClientWidth     =   11610
   FillColor       =   &H00004000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   11610
   Begin VB.CommandButton cmdQuit 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   4200
      Picture         =   "frmStart.frx":0000
      ScaleHeight     =   2715
      ScaleWidth      =   2955
      TabIndex        =   2
      Top             =   3000
      Width           =   3015
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00008000&
      Caption         =   "Start"
      Height          =   1095
      Left            =   8160
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5880
      Width           =   3135
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Minnesota Wild Hockey "
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1455
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   10695
   End
End
Attribute VB_Name = "frmMinnesotaWildHockeyTeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : MinnesotaWildTeamInformationProgram (Stephanie Weimar's VB Project.vbp)
'Form Name : frmMinnesotaWildHockeyTeam (frmMinnesotaWildHockeyTeam.frm)
'Author: Stephanie Weimar
'Date Written: October 29, 2003
'Purpose of Form: To give user option of begining program or ending

'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Private Sub cmdQuit_Click()             'Use this button to quit a program
    End
End Sub
Private Sub cmdStart_Click()                    'This takes you from the first form the the second
    frmMinnesotaWildHockeyTeam.Visible = False
    frmRoster.Visible = True
End Sub
