VERSION 5.00
Begin VB.Form FrmCuddyer 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   9615
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
      Left            =   8520
      Picture         =   "FrmCuddyer.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6720
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
      Left            =   6600
      Picture         =   "FrmCuddyer.frx":16CE
      Style           =   1  'Graphical
      TabIndex        =   12
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
      Picture         =   "FrmCuddyer.frx":2D9C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6600
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Index           =   1
      Left            =   0
      Picture         =   "FrmCuddyer.frx":446A
      ScaleHeight     =   1695
      ScaleWidth      =   1695
      TabIndex        =   4
      Top             =   0
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   240
      Picture         =   "FrmCuddyer.frx":E334
      ScaleHeight     =   5655
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   1680
      Width           =   4335
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   $"FrmCuddyer.frx":7C092
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
      Left            =   5160
      TabIndex        =   10
      Top             =   5160
      Width           =   4095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Height: 6'2""           Weight: 202 lbs"
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
      Left            =   5280
      TabIndex        =   9
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
      Left            =   5160
      TabIndex        =   8
      Top             =   4200
      Width           =   3975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hometown: Norfolk, VA"
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
      Left            =   5160
      TabIndex        =   7
      Top             =   3720
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Birthdate: March 27, 1979"
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
      Left            =   5640
      TabIndex        =   6
      Top             =   3240
      Width           =   3735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name: Michael Brent Cuddyer"
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
      Left            =   4800
      TabIndex        =   5
      Top             =   2760
      Width           =   4695
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000012&
      Caption         =   "#5"
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
      Left            =   8640
      TabIndex        =   3
      Top             =   360
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
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Michael Cuddyer"
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
      Height          =   2535
      Left            =   4680
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "FrmCuddyer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: Minnesota Twins Fan
'Form Name: FrmCuddyer
'Project By: Stephanie Arel
'Date Written: 3/12/2009
'This Form just lists the facts about Michael Cuddyer.
Option Explicit

Private Sub CmdMain_Click()
'Takes the user back to the main menu
FrmCuddyer.Hide
FrmMain.Show
End Sub

Private Sub CmdQuit_Click()
'Ends the Program
End
End Sub

Private Sub CmdReturn_Click()
'Takes User back to the main Players menu
FrmCuddyer.Hide
FrmPlayers.Show
End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
