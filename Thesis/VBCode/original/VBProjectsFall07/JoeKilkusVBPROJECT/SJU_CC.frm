VERSION 5.00
Begin VB.Form frmSJU_CC 
   Caption         =   "Form2"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
   LinkTopic       =   "Form2"
   Picture         =   "SJU_CC.frx":0000
   ScaleHeight     =   7035
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWorksCited 
      Caption         =   "Works Cited"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5640
      TabIndex        =   6
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7320
      TabIndex        =   5
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdResults 
      Caption         =   "2007 Results"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3960
      TabIndex        =   4
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdRoster 
      Caption         =   "2007 Roster"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2280
      TabIndex        =   3
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdTeamInformation 
      Caption         =   "Team Information"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      TabIndex        =   2
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label lbl2007 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "2007"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   3600
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label lblSJU_CC 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Saint John's Cross Country"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "frmSJU_CC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'the purpose of this program is to give the user a virtual tour
'of the history of Saint John's Cross Country, and give them
'a look at the 2007 team


'this button ends the program
Private Sub cmdQuit_Click()
    End
End Sub

'this button takes the user to the results screen
Private Sub cmdResults_Click()
    frmSJU_CC.Hide
    frmSeason_Results.Show
    cmdResults.Enabled = False
    cmdWorksCited.Enabled = True
End Sub

'this button takes the user to the roster screen
Private Sub cmdRoster_Click()
    frmSJU_CC.Hide
    frmRoster.Show
    cmdRoster.Enabled = False
    cmdResults.Enabled = True
End Sub

'this button takes the user to the team information page
Private Sub cmdTeamInformation_Click()
    frmSJU_CC.Hide
    frmTeam_Info.Show
    cmdTeamInformation.Enabled = False
    cmdRoster.Enabled = True
End Sub

'this button takes the user to the Works Cited page
Private Sub cmdWorksCited_Click()
    frmSJU_CC.Hide
    frmWorksCited.Show
    cmdWorksCited.Enabled = False
    cmdQuit.Enabled = True
End Sub
