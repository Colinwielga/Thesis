VERSION 5.00
Begin VB.Form FrmGomez 
   BackColor       =   &H80000012&
   Caption         =   "Form1"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   10020
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
      Left            =   3600
      Picture         =   "FrmGomez.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6720
      Width           =   1335
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
      Left            =   1680
      Picture         =   "FrmGomez.frx":16CE
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6720
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
      Left            =   360
      Picture         =   "FrmGomez.frx":2D9C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6840
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   5415
      Left            =   5520
      Picture         =   "FrmGomez.frx":446A
      ScaleHeight     =   5355
      ScaleWidth      =   3795
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H8000000E&
      Height          =   5655
      Index           =   0
      Left            =   5400
      ScaleHeight     =   5595
      ScaleWidth      =   4035
      TabIndex        =   1
      Top             =   360
      Width           =   4095
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00800000&
      Height          =   5895
      Left            =   5280
      ScaleHeight     =   5835
      ScaleWidth      =   4275
      TabIndex        =   2
      Top             =   240
      Width           =   4335
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Index           =   1
      Left            =   8400
      Picture         =   "FrmGomez.frx":610AC
      ScaleHeight     =   1695
      ScaleWidth      =   1695
      TabIndex        =   4
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "* Twins manager Ron Gardenhire has nicknamed him ""Go-Go"" due to his last name and his blazing speed."
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
      Height          =   1335
      Left            =   960
      TabIndex        =   15
      Top             =   5400
      Width           =   4095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Height: 6'4""           Weight: 215 lbs"
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
      Left            =   840
      TabIndex        =   14
      Top             =   4680
      Width           =   3735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Bats: Right             Throws: Right"
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
      Left            =   600
      TabIndex        =   13
      Top             =   4200
      Width           =   3975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hometown: Santiago, Dominican Republic "
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
      Left            =   600
      TabIndex        =   12
      Top             =   3600
      Width           =   4215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Birthdate: December 4, 1985"
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
      Left            =   960
      TabIndex        =   11
      Top             =   3120
      Width           =   3735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name: Carlos Argelis Gomez"
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
      Left            =   360
      TabIndex        =   10
      Top             =   2640
      Width           =   4695
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
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Carlos Gomez"
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
      Height          =   2295
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000012&
      Caption         =   "#22"
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
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "FrmGomez"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: Minnesota Twins Fan
'Form Name: FrmGomez
'Project By: Stephanie Arel
'Date Written: 3/12/2009
'This Form just lists the facts about Carlos Gomez.
Option Explicit

Private Sub CmdMain_Click()
'Takes User Back to Main Menu
FrmGomez.Hide
FrmMain.Show
End Sub

Private Sub CmdQuit_Click()
'Ends Program Completely
End
End Sub

Private Sub CmdReturn_Click()
'Takes User back to main players menu
FrmGomez.Hide
FrmPlayers.Show
End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
