VERSION 5.00
Begin VB.Form FrmCasilla 
   BackColor       =   &H80000012&
   Caption         =   "Form1"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows Default
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
      Left            =   240
      Picture         =   "FrmCasilla.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6960
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
      Left            =   1320
      Picture         =   "FrmCasilla.frx":16CE
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6840
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
      Left            =   2880
      Picture         =   "FrmCasilla.frx":2D9C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6840
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   840
      Picture         =   "FrmCasilla.frx":446A
      ScaleHeight     =   4545
      ScaleWidth      =   3345
      TabIndex        =   0
      Top             =   2160
      Width           =   3375
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      ForeColor       =   &H80000008&
      Height          =   6135
      Index           =   0
      Left            =   360
      ScaleHeight     =   6105
      ScaleWidth      =   9225
      TabIndex        =   1
      Top             =   240
      Width           =   9255
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4455
         Left            =   120
         ScaleHeight     =   4455
         ScaleWidth      =   4095
         TabIndex        =   10
         Top             =   1680
         Width           =   4095
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmCasilla.frx":37644
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1335
         Left            =   4560
         TabIndex        =   11
         Top             =   4680
         Width           =   4095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name: Alexi Lora Casilla "
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
         Left            =   4320
         TabIndex        =   9
         Top             =   1800
         Width           =   4695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Birthdate: July 20, 1984"
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
         Left            =   5280
         TabIndex        =   8
         Top             =   2280
         Width           =   3735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hometown: San Cristobal, Dominican Republic"
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
         Left            =   4800
         TabIndex        =   7
         Top             =   2760
         Width           =   3975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Bats: Both            Throws: Right"
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
         Left            =   4680
         TabIndex        =   6
         Top             =   3600
         Width           =   3975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Height: 5'9""           Weight: 160 lbs"
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
         Left            =   4800
         TabIndex        =   5
         Top             =   4080
         Width           =   3735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2nd Base"
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
         TabIndex        =   4
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "#25"
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
         Left            =   480
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Alexi Casilla"
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
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   6735
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Index           =   1
      Left            =   8400
      Picture         =   "FrmCasilla.frx":376DC
      ScaleHeight     =   1695
      ScaleWidth      =   1695
      TabIndex        =   12
      Top             =   6240
      Width           =   1695
   End
End
Attribute VB_Name = "FrmCasilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: Minnesota Twins Fan
'Form Name: FrmCasilla
'Project By: Stephanie Arel
'Date Written: 3/12/2009
'This Form just lists the facts about Alexi Casilla.
Option Explicit

Private Sub CmdMain_Click()
'Takes the User back to the Main Menu
FrmCasilla.Hide
FrmMain.Show
End Sub

Private Sub CmdQuit_Click()
'Ends the Program Completely
End
End Sub

Private Sub CmdReturn_Click()
'Takes the User back to the Players Menu
FrmCasilla.Hide
FrmPlayers.Show
End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
