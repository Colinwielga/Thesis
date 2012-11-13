VERSION 5.00
Begin VB.Form FrmYoung 
   Caption         =   "Form1"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7695
      Left            =   0
      Picture         =   "FrmYoung.frx":0000
      ScaleHeight     =   7695
      ScaleWidth      =   9375
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Index           =   1
         Left            =   120
         Picture         =   "FrmYoung.frx":BE8D2
         ScaleHeight     =   1695
         ScaleWidth      =   1695
         TabIndex        =   11
         Top             =   6000
         Width           =   1695
      End
      Begin VB.CommandButton CmdMain 
         BackColor       =   &H8000000E&
         Caption         =   "Go Back to Main Menu"
         BeginProperty Font 
            Name            =   "Rockwell Extra Bold"
            Size            =   9
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6480
         Picture         =   "FrmYoung.frx":C879C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6600
         Width           =   1335
      End
      Begin VB.CommandButton CmdReturn 
         BackColor       =   &H8000000E&
         Caption         =   "Return to Players"
         BeginProperty Font 
            Name            =   "Rockwell Extra Bold"
            Size            =   9
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4800
         Picture         =   "FrmYoung.frx":C9E6A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   6600
         Width           =   1335
      End
      Begin VB.CommandButton CmdQuit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Rockwell Extra Bold"
            Size            =   9
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8160
         Picture         =   "FrmYoung.frx":CB538
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   6720
         Width           =   735
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   0
         ScaleHeight     =   1695
         ScaleWidth      =   9255
         TabIndex        =   12
         Top             =   6000
         Width           =   9255
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "* Delmon's brother, Dmitri, plays for the St. Louis Cardinals."
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   4320
         TabIndex        =   15
         Top             =   5280
         Width           =   4095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Height: 6'3""           Weight: 205 lbs"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   3000
         Width           =   3735
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Bats: Right             Throws: Right"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   735
         Left            =   0
         TabIndex        =   10
         Top             =   2640
         Width           =   3975
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hometown: Montgomery, AL"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   3975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Birthdate: September 14, 1985"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   1680
         Width           =   3735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name: Delmon Damarcus Young "
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   -360
         TabIndex        =   7
         Top             =   1200
         Width           =   5415
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "#21"
         BeginProperty Font 
            Name            =   "Rockwell Extra Bold"
            Size            =   12
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Outfield"
         BeginProperty Font 
            Name            =   "Rockwell Condensed"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   6240
         TabIndex        =   2
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Delmon Young"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   840
         TabIndex        =   1
         Top             =   120
         Width           =   6735
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Height: 6'4""           Weight: 235 lbs"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "FrmYoung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: Minnesota Twins Fan
'Form Name: FrmYoung
'Project By: Stephanie Arel
'Date Written: 3/12/2009
'This Form just lists the facts about Delmon Young.
Option Explicit


Private Sub CmdMain_Click()
'Takes user back to main menu
FrmYoung.Hide
FrmMain.Show
End Sub

Private Sub CmdQuit_Click()
'Ends Program
End
End Sub

Private Sub CmdReturn_Click()
'Takes user back to players menu
FrmYoung.Hide
FrmPlayers.Show
End Sub


Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
