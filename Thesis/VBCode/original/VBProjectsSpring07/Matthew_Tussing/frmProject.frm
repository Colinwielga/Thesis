VERSION 5.00
Begin VB.Form frmProject 
   BackColor       =   &H00FF0000&
   Caption         =   "Project"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWorks 
      Caption         =   "Go to the WORKS CITED page"
      Height          =   975
      Left            =   2400
      TabIndex        =   4
      Top             =   6240
      Width           =   5055
   End
   Begin VB.CommandButton cmdNL 
      Caption         =   "Go to the National Leaugs Page for stats from 2006"
      Height          =   2415
      Left            =   1200
      TabIndex        =   3
      Top             =   3360
      Width           =   3255
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   2415
      Left            =   5280
      TabIndex        =   2
      Top             =   3360
      Width           =   3255
   End
   Begin VB.CommandButton cmdAL 
      Caption         =   "Go to the American League Page for stats from 2006"
      Height          =   2535
      Left            =   1200
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.CommandButton cmdTwins 
      Caption         =   "Go to the Minnesota Twins Page"
      Height          =   2535
      Left            =   5280
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "frmProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAL_Click()
    frmAL.Show 'goes to the AL screen
    frmProject.Hide
End Sub


Private Sub cmdNL_Click()
    frmNL.Show 'goes to the NL screen
    frmProject.Hide
End Sub

Private Sub cmdQuit_Click()
End 'ends the program
End Sub

Private Sub cmdTwins_Click()
    frmTwins.Show 'goes to the twins screen
    frmProject.Hide
End Sub

Private Sub cmdWorks_Click()
    frmWorks.Show 'goes to the works cited screen
    frmProject.Hide
End Sub
