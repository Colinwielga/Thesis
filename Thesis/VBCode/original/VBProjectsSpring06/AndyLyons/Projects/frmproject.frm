VERSION 5.00
Begin VB.Form frmNFLDraft 
   BackColor       =   &H8000000D&
   Caption         =   "Draft Busts? or MVP's?"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10485
   FillColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   Picture         =   "frmproject.frx":0000
   ScaleHeight     =   7800
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDisplay 
      Height          =   615
      Left            =   2880
      ScaleHeight     =   555
      ScaleWidth      =   3195
      TabIndex        =   6
      Top             =   3360
      Width           =   3255
   End
   Begin VB.CommandButton cmdchoice 
      Caption         =   "Click to Enter Name"
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   2760
      Width           =   3255
   End
   Begin VB.CommandButton cmddefprofiles 
      Caption         =   "View Defensive Player Profiles"
      Height          =   975
      Left            =   4680
      TabIndex        =   4
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdoffprofiles 
      Caption         =   "View Offensive Player Profiles"
      Height          =   975
      Left            =   2880
      TabIndex        =   3
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search Players Stats"
      Height          =   1095
      Left            =   2880
      TabIndex        =   2
      Top             =   4080
      Width           =   3255
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdDraft 
      Caption         =   "Draft Central!! Pick Your Player after searching stats and profiles!"
      Height          =   975
      Left            =   3000
      TabIndex        =   0
      Top             =   6480
      Width           =   3015
   End
   Begin VB.Label lblfantasy 
      BackColor       =   &H0000FFFF&
      Caption         =   "Your Own Fantasy Draft! Search the Profiles, stats, Choose a team, and fill your teams need for the 2006 season!"
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2760
      TabIndex        =   7
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Image Image4 
      Height          =   2415
      Left            =   0
      Picture         =   "frmproject.frx":21462
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Image Image3 
      Height          =   2985
      Left            =   6360
      Picture         =   "frmproject.frx":335F0
      Top             =   4440
      Width           =   2850
   End
   Begin VB.Image Image2 
      Height          =   2985
      Left            =   120
      Picture         =   "frmproject.frx":4F2D6
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   6360
      Picture         =   "frmproject.frx":64604
      Top             =   1200
      Width           =   2745
   End
End
Attribute VB_Name = "frmNFLDraft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim A As String
Dim B As String
Dim C As String
Dim Size As Integer
Dim Pos As Integer
'2006 NFL Draft Simulator (Draft.vbp)
'frmNFLDraft(frmproject.frm)
'Andy Lyons
'March 24, 2006
'Uploads Beginning options where user can click on buttons to view statistics and simulate a draft
'The purpose of this project is to give the user the ability to choose who they think will be the best athlete for their team by looking at profiles and data.
'clicking this allows the user to input their name, and be refered to as the coach of the team.
Private Sub cmdchoice_Click()
    MsgBox "Welcome to the 2006 NFL Fantasy Draft", , "Welcome"
    A = InputBox("What is your first name?", "First Name")
    B = InputBox("What is your last name?", "Last Name")
        picDisplay.Print "Coach " & B
End Sub
'Clicking this button brings the user to the Defensive Player Profiles, where user can use more options.
Private Sub cmddefprofiles_Click()
    frmdefpositions.Show
End Sub
'Clicking this button brings the user to the Draft Simulator, where the user can choose their team and player.
Private Sub cmdDraft_Click()
    frmdraft.Show
End Sub

Private Sub cmdExit_Click()
    End
End Sub

'Clicking this button brings the user to the Offensive Player Profiles, where the user can use more options.
Private Sub cmdoffprofiles_Click()
    frmoffpositions.Show
End Sub
'Clicking this button brings you to a setup where you can look up stats of all the players eligible. In this function you can see lists of who has the highests bench press, vertical leap, and forty yard dash.
Private Sub cmdsearch_Click()
    frmstats.Show
       
End Sub

