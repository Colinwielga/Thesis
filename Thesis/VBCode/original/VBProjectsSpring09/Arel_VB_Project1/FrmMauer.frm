VERSION 5.00
Begin VB.Form FrmMauer 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   Picture         =   "FrmMauer.frx":0000
   ScaleHeight     =   6840
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
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
      Left            =   5760
      Picture         =   "FrmMauer.frx":3D022
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5880
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
      Left            =   9120
      Picture         =   "FrmMauer.frx":3E6F0
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6000
      Width           =   735
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
      Left            =   7320
      Picture         =   "FrmMauer.frx":3FDBE
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5880
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      Picture         =   "FrmMauer.frx":4148C
      ScaleHeight     =   1695
      ScaleWidth      =   1695
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   960
      Picture         =   "FrmMauer.frx":4B356
      ScaleHeight     =   4095
      ScaleWidth      =   3855
      TabIndex        =   0
      Top             =   2160
      Width           =   3855
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   600
      ScaleHeight     =   4785
      ScaleWidth      =   4545
      TabIndex        =   2
      Top             =   1800
      Width           =   4575
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   0
      ScaleHeight     =   5625
      ScaleWidth      =   10185
      TabIndex        =   3
      Top             =   0
      Width           =   10215
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         Caption         =   "* Joe struck out just once in his high school career and finished with a .567 average. "
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   855
         Left            =   5760
         TabIndex        =   13
         Top             =   4800
         Width           =   3735
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         Caption         =   "* Joe was named All-State in three sports during his senior year at Cretin-Derham. "
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1095
         Left            =   5640
         TabIndex        =   12
         Top             =   3720
         Width           =   4095
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000012&
         Caption         =   "#7"
         BeginProperty Font 
            Name            =   "Rockwell Extra Bold"
            Size            =   12
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   9240
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Height: 6'5""           Weight: 220 lbs"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   5760
         TabIndex        =   10
         Top             =   3120
         Width           =   3735
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         Caption         =   "Bats: Left              Throws: Right"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   735
         Left            =   5520
         TabIndex        =   9
         Top             =   2760
         Width           =   3975
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hometown: St. Paul, MN"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   5400
         TabIndex        =   8
         Top             =   2280
         Width           =   3975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Birthdate: April 19, 1983 "
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   6000
         TabIndex        =   7
         Top             =   1920
         Width           =   3735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name: Joseph Patrick Mauer"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   5400
         TabIndex        =   6
         Top             =   1560
         Width           =   4095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Catcher"
         BeginProperty Font 
            Name            =   "Rockwell Condensed"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   1200
         TabIndex        =   5
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "JOE MAUER"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1215
         Left            =   3240
         TabIndex        =   4
         Top             =   360
         Width           =   6375
      End
   End
End
Attribute VB_Name = "FrmMauer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: Minnesota Twins Fan
'Form Name: FrmMauer
'Project By: Stephanie Arel
'Date Written: 3/12/2009
'This Form just lists the facts about Joe Mauer.
Option Explicit

Private Sub CmdMain_Click()
'Takes the user back to the main menu
FrmMauer.Hide
FrmMain.Show
End Sub

Private Sub CmdQuit_Click()
'Ends the program
End
End Sub

Private Sub CmdReturn_Click()
'Takes the user back to the main players menu
FrmMauer.Hide
FrmPlayers.Show
End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
