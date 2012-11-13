VERSION 5.00
Begin VB.Form FrmSpan 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   7575
      Index           =   0
      Left            =   480
      Picture         =   "FrmSpan.frx":0000
      ScaleHeight     =   7545
      ScaleWidth      =   10305
      TabIndex        =   0
      Top             =   120
      Width           =   10335
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
         Picture         =   "FrmSpan.frx":C6102
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   6840
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
         Left            =   7680
         Picture         =   "FrmSpan.frx":C77D0
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   6720
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
         Left            =   6240
         Picture         =   "FrmSpan.frx":C8E9E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   6720
         Width           =   1335
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Index           =   1
         Left            =   8400
         Picture         =   "FrmSpan.frx":CA56C
         ScaleHeight     =   1695
         ScaleWidth      =   1695
         TabIndex        =   10
         Top             =   -120
         Width           =   1695
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "*On July 28, 2008, Span hit his first major league career home run off Mark Buehrle of the Chicago White Sox."
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1575
         Left            =   6840
         TabIndex        =   9
         Top             =   5040
         Width           =   3135
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name: Keiunta Denard Span "
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   6840
         TabIndex        =   8
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Birthdate: February 27, 1984 "
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   7080
         TabIndex        =   7
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hometown: Tampa, Florida"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   6840
         TabIndex        =   6
         Top             =   2880
         Width           =   3255
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Bats: Left Throws: Left"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   7560
         TabIndex        =   5
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Height: 6'0"" Weight: 205 lbs"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   7320
         TabIndex        =   4
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "#2"
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
         Top             =   0
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
         Left            =   1800
         TabIndex        =   2
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Denard Span"
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
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   6735
      End
   End
End
Attribute VB_Name = "FrmSpan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: Minnesota Twins Fan
'Form Name: FrmSpan
'Project By: Stephanie Arel
'Date Written: 3/12/2009
'This Form just lists the facts about Denard Span.
Option Explicit


Private Sub CmdMain_Click()
'Takes user back to main menu
FrmSpan.Hide
FrmMain.Show
End Sub

Private Sub CmdQuit_Click()
'Ends Program
End
End Sub

Private Sub CmdReturn_Click()
'Takes User Back to Main Player Menu
FrmSpan.Hide
FrmPlayers.Show
End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
