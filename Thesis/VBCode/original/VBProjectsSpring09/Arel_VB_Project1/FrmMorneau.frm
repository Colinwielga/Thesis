VERSION 5.00
Begin VB.Form FrmMorneau 
   BackColor       =   &H80000012&
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   10275
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
      Picture         =   "FrmMorneau.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6600
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
      Left            =   1440
      Picture         =   "FrmMorneau.frx":16CE
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6480
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
      Left            =   3240
      Picture         =   "FrmMorneau.frx":2D9C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6480
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   0
      Left            =   5160
      Picture         =   "FrmMorneau.frx":446A
      ScaleHeight     =   4695
      ScaleWidth      =   4575
      TabIndex        =   1
      Top             =   2280
      Width           =   4575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   360
      ScaleHeight     =   2655
      ScaleWidth      =   5415
      TabIndex        =   5
      Top             =   1560
      Width           =   5415
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
         Left            =   840
         TabIndex        =   10
         Top             =   2160
         Width           =   3735
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000080&
         Height          =   735
         Left            =   600
         TabIndex        =   9
         Top             =   1800
         Width           =   3975
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hometown: New Westminster, British Columbia, Canada"
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
         Height          =   735
         Left            =   600
         TabIndex        =   8
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Birthdate: May 15, 1981"
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
         Height          =   495
         Left            =   1200
         TabIndex        =   7
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name: Justin Ernest George Morneau "
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
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   5415
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Index           =   1
      Left            =   8640
      Picture         =   "FrmMorneau.frx":5DD54
      ScaleHeight     =   1695
      ScaleWidth      =   1695
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "* In 2008, he became the first Canadian to win the Home Run Derby."
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
      Left            =   960
      TabIndex        =   12
      Top             =   5400
      Width           =   3735
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "* Justin was awarded the American League Most Valuable Player award in 2006."
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
      Left            =   720
      TabIndex        =   11
      Top             =   4440
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Justin Morneau"
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
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   6735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1st Base"
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
      Left            =   5640
      TabIndex        =   3
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000012&
      Caption         =   "#33"
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
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "FrmMorneau"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: Minnesota Twins Fan
'Form Name: FrmMorneau
'Project By: Stephanie Arel
'Date Written: 3/12/2009
'This Form just lists the facts about Justin Morneau.
Option Explicit

Private Sub CmdMain_Click()
'Takes user back to main menu
FrmMorneau.Hide
FrmMain.Show
End Sub

Private Sub CmdQuit_Click()
'Ends the program
End
End Sub

Private Sub CmdReturn_Click()
'Takes user back to the main Players menu
FrmMorneau.Hide
FrmPlayers.Show
End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
