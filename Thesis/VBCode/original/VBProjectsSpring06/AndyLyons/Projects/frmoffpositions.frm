VERSION 5.00
Begin VB.Form frmoffpositions 
   BackColor       =   &H0000FFFF&
   Caption         =   "Positions"
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   ScaleHeight     =   9075
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      BackColor       =   &H000000FF&
      Caption         =   "Back to Main Menu"
      Height          =   855
      Left            =   5280
      MaskColor       =   &H000000FF&
      TabIndex        =   8
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label lblrunning 
      Caption         =   "Running Backs"
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lbltackles 
      Caption         =   "Tackles"
      Height          =   255
      Left            =   7200
      TabIndex        =   7
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label lblguards 
      Caption         =   "Guards"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label lblte 
      Caption         =   "Tight Ends"
      Height          =   255
      Left            =   7560
      TabIndex        =   5
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lblreceivers 
      Caption         =   "Wide Receivers"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lblblank 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Offensive Player Profiles"
      BeginProperty Font 
         Name            =   "Niagara Engraved"
         Size            =   39.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1815
      Left            =   3600
      TabIndex        =   3
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Label lblcenters 
      Caption         =   "Centers"
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Image imgrbs 
      Height          =   3165
      Left            =   360
      Picture         =   "frmoffpositions.frx":0000
      Top             =   600
      Width           =   2655
   End
   Begin VB.Image imcts 
      Height          =   2250
      Left            =   4080
      Picture         =   "frmoffpositions.frx":1B6BE
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label lblquaterbacks 
      Caption         =   "Quaterbacks"
      Height          =   255
      Left            =   7440
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Image imggrds 
      Height          =   1950
      Left            =   360
      Picture         =   "frmoffpositions.frx":27028
      Top             =   6960
      Width           =   2850
   End
   Begin VB.Image imgtcks 
      Height          =   2565
      Left            =   6720
      Picture         =   "frmoffpositions.frx":392E2
      Top             =   6360
      Width           =   2700
   End
   Begin VB.Label Label1 
      Caption         =   "Fullbacks"
      Height          =   255
      Left            =   4320
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
   Begin VB.Image imgfbs 
      Height          =   2625
      Left            =   3720
      Picture         =   "frmoffpositions.frx":4FBD8
      Top             =   720
      Width           =   2100
   End
   Begin VB.Image imgtes 
      Height          =   1995
      Left            =   6840
      Picture         =   "frmoffpositions.frx":61B36
      Top             =   3720
      Width           =   2475
   End
   Begin VB.Image imgwrs 
      Height          =   2325
      Left            =   360
      Picture         =   "frmoffpositions.frx":71D28
      Top             =   4200
      Width           =   2790
   End
   Begin VB.Image imgqbs 
      Height          =   2460
      Left            =   6480
      Picture         =   "frmoffpositions.frx":8707A
      Top             =   840
      Width           =   2985
   End
End
Attribute VB_Name = "frmoffpositions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'2006 NFL Draft Simulator (Draft.vbp)
'Offensive Positions(frmoffpositions)
'Andy Lyons
'March 24, 2006
'Uploads Beginning options where user can click on buttons to view statistics and simulate a draft
'The purpose of this project is to give the user the ability to choose who they think will be the best athlete for their team by looking at profiles and data.

Private Sub cmdback_Click()
'this button allows you to go back to the Main startup screen with options
    frmNFLDraft.Show
    frmoffpositions.Hide
End Sub

Private Sub imgrbs_Click()
'Clicking this image brings you to the Runningback menu to see their profiles.
    frmrunningbacks.Show
End Sub
'Clicking this button brings the user to the Center
Private Sub imcts_Click()
    frmcenters.Show
End Sub

Private Sub imgfbs_Click()
    frmfbs.Show
End Sub

Private Sub imggrds_Click()
    frmguards.Show
End Sub

Private Sub imgqbs_Click()
    frmqbs.Show
End Sub

Private Sub imgtcks_Click()
    frmtackles.Show
End Sub

Private Sub imgtes_Click()
    frmtes.Show
End Sub

Private Sub imgwrs_Click()
    frmwrs.Show
End Sub
