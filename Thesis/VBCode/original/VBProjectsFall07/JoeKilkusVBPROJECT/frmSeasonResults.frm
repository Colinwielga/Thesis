VERSION 5.00
Begin VB.Form frmSeason_Results 
   BackColor       =   &H80000013&
   Caption         =   "Form1"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBackToMain 
      Caption         =   "Back to Front Page"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   10
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton cmdViewSchedule 
      Caption         =   "View Season Schedule"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   720
      Width           =   2295
   End
   Begin VB.PictureBox picSchedule 
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
      Height          =   2775
      Left            =   360
      ScaleHeight     =   2715
      ScaleWidth      =   8715
      TabIndex        =   7
      Top             =   1200
      Width           =   8775
   End
   Begin VB.CommandButton cmdViewMIAC 
      Caption         =   "MIAC Championship"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7920
      TabIndex        =   6
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdViewJimDrews 
      Caption         =   "Jim Drews Invite"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6360
      TabIndex        =   5
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdViewPreNats 
      Caption         =   "Pre-Nationals"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4800
      TabIndex        =   4
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdViewWillamette 
      Caption         =   "Willamette Invite"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3240
      TabIndex        =   3
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdViewEauClaire 
      Caption         =   "UW-Eau Claire Invite"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      TabIndex        =   2
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdViewSJUInvite 
      Caption         =   "SJU Invite"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "View Results from . . ."
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   2880
      TabIndex        =   9
      Top             =   4080
      Width           =   3975
   End
   Begin VB.Label lblSJU_CC 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "2007 Schedule/Results "
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
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   9135
   End
End
Attribute VB_Name = "frmSeason_Results"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'pressing this button takes the user back to the front page
Private Sub cmdBacktoMain_Click()
    frmSeason_Results.Hide
    frmSJU_CC.Show
End Sub

'pressing this button takes the user to the results from this race,
'and allows them to view the results from the next race in the list
'when they return to this form
Private Sub cmdViewEauClaire_Click()
    frmSeason_Results.Hide
    frmEau_Claire_Results.Show
    cmdViewEauClaire.Enabled = False
    cmdViewWillamette.Enabled = True
End Sub

'pressing this button takes the user to the results from this race,
'and allows them to view the results from the next race in the list
'when they return to this form
Private Sub cmdViewJimDrews_Click()
    frmSeason_Results.Hide
    frmJim_Drews_Results.Show
    cmdViewJimDrews.Enabled = False
    cmdViewMIAC.Enabled = True
End Sub

'pressing this button takes the user to the results from this race,
'and allows them to view the results from the next race in the list
'when they return to this form
Private Sub cmdViewMIAC_Click()
    frmSeason_Results.Hide
    frmMIAC_Results.Show
    cmdViewMIAC.Enabled = False
    cmdBacktoMain.Enabled = True
End Sub

'pressing this button takes the user to the results from this race,
'and allows them to view the results from the next race in the list
'when they return to this form
Private Sub cmdViewPreNats_Click()
    frmSeason_Results.Hide
    frmPreNats_Results.Show
    cmdViewPreNats.Enabled = False
    cmdViewJimDrews.Enabled = True
End Sub

'pressing this button displays the 2007 SJU Cross Country schedule,
'and how the team finished at each meet
Private Sub cmdViewSchedule_Click()
    picSchedule.Cls
    picSchedule.Print "Date"; Tab(15); "Event"; Tab(38); "Site"; Tab(59); "Place"
    picSchedule.Print "--------------------------------------------------------------------------------------------------------------------------"
    picSchedule.Print "9/8/2007"; Tab(15); "SJU Invite"; Tab(38); "Collegeville, MN"; Tab(59); "T-1 out of 10"
    picSchedule.Print "9/21/2007"; Tab(15); "UW-Eau Claire Invite"; Tab(38); "Colfax, WI"; Tab(59); "11 out of 20"
    picSchedule.Print "9/29/2007"; Tab(15); "Willamette Invite"; Tab(38); "Salem, OR"; Tab(59); "3 out of 27"
    picSchedule.Print "10/6/2007"; Tab(15); "Pre-Nationals"; Tab(38); "Northfield, MN"; Tab(59); "9 out of 15"
    picSchedule.Print "10/13/2007"; Tab(15); "Jim Drews Invite"; Tab(38); "West Salem, WI"; Tab(59); "4 out of 25"
    picSchedule.Print "10/27/2007"; Tab(15); "MIAC Championship"; Tab(38); "St. Paul, MN"; Tab(59); "1 out of 11"
    cmdViewSJUInvite.Enabled = True
    cmdViewSchedule.Enabled = False
End Sub

'pressing this button takes the user to the results from this race,
'and allows them to view the results from the next race in the list
'when they return to this form
Private Sub cmdViewSJUInvite_Click()
    frmSeason_Results.Hide
    frmSJU_Invite_Results.Show
    cmdViewSJUInvite.Enabled = False
    cmdViewEauClaire.Enabled = True
End Sub

'pressing this button takes the user to the results from this race,
'and allows them to view the results from the next race in the list
'when they return to this form
Private Sub cmdViewWillamette_Click()
    frmSeason_Results.Hide
    frmWillamette_Results.Show
    cmdViewWillamette.Enabled = False
    cmdViewPreNats.Enabled = True
End Sub
